import fs from 'fs';
import * as path from 'path';
import os from 'os';
import { spawn } from 'child_process';
import * as msal from '@azure/msal-node';
import { login, logout, getAccessToken, getAccount, isAuthenticated, listProfiles } from '../src/auth/authManager';
import { withProfile } from '../src/profile';

jest.mock('@azure/msal-node', () => ({ PublicClientApplication: jest.fn() }));

const alice = {
  homeAccountId: 'alice-home', environment: 'login.microsoftonline.com', tenantId: 'tenant',
  username: 'alice@example.com', name: 'Alice', localAccountId: 'alice-local',
};
const bob = { ...alice, homeAccountId: 'bob-home', username: 'bob@example.com', name: 'Bob' };
let home: string;
let selected = alice;
let failLogin = false;
let failSilent = false;
let silentAccount: typeof alice | undefined;
let previousClientId: string | undefined;

function file(profile = 'default'): string {
  return path.join(home, '.ms-todo-cli', 'profiles', `${profile}.json`);
}

function read(profile = 'default'): { account: typeof alice; cache: string } {
  return JSON.parse(fs.readFileSync(file(profile), 'utf8'));
}

beforeEach(() => {
  home = fs.mkdtempSync(path.join(process.cwd(), '.auth-test-'));
  jest.spyOn(os, 'homedir').mockReturnValue(home);
  previousClientId = process.env.MS_TODO_CLIENT_ID;
  process.env.MS_TODO_CLIENT_ID = 'test-client';
  selected = alice;
  failLogin = false;
  failSilent = false;
  silentAccount = undefined;
  (msal.PublicClientApplication as jest.Mock).mockImplementation(() => {
    let state = { accounts: [] as typeof alice[], revision: 0 };
    const cache = {
      serialize: () => JSON.stringify(state),
      deserialize: (serialized: string) => { state = JSON.parse(serialized); },
      getAllAccounts: async () => state.accounts,
    };
    return {
      getTokenCache: () => cache,
      acquireTokenByDeviceCode: async () => {
        state.accounts = [selected];
        if (failLogin) throw new Error('Login cancelled');
        return { account: selected, accessToken: 'test-access-token' };
      },
      acquireTokenSilent: async ({ account }: { account: typeof alice }) => {
        await new Promise(resolve => setTimeout(resolve, 5));
        state.revision += 1;
        if (failSilent) throw new Error('Refresh failed');
        return { account: silentAccount || account, accessToken: 'test-access-token' };
      },
    };
  });
});

afterEach(() => {
  jest.restoreAllMocks();
  if (previousClientId === undefined) delete process.env.MS_TODO_CLIENT_ID;
  else process.env.MS_TODO_CLIENT_ID = previousClientId;
  fs.rmSync(home, { recursive: true, force: true });
});

test('profiles are isolated and metadata/listing do not require a client ID or expose tokens', async () => {
  await withProfile('work', () => login());
  selected = bob;
  await withProfile('personal', () => login());
  delete process.env.MS_TODO_CLIENT_ID;
  expect(await withProfile('work', () => getAccount())).toMatchObject({ homeAccountId: alice.homeAccountId });
  expect(await withProfile('missing', () => isAuthenticated())).toBe(false);
  const profiles = await listProfiles();
  expect(profiles.map(entry => entry.profile)).toEqual(['default', 'personal', 'work']);
  expect(JSON.stringify(profiles)).not.toContain('test-access-token');
  expect(JSON.stringify(profiles)).not.toContain('"cache"');
  await withProfile('work', () => logout());
  expect(await withProfile('work', () => getAccount())).toBeNull();
  expect(await withProfile('personal', () => isAuthenticated())).toBe(true);
});

test('different stable identities require replacement and rejected/failed logins preserve disk', async () => {
  await login();
  const original = fs.readFileSync(file(), 'utf8');
  selected = { ...alice, username: 'renamed@example.com' };
  await login();
  expect((await getAccount())?.username).toBe('renamed@example.com');
  const renamed = fs.readFileSync(file(), 'utf8');
  selected = bob;
  await expect(login()).rejects.toThrow('--replace');
  expect(fs.readFileSync(file(), 'utf8')).toBe(renamed);
  failLogin = true;
  await expect(login({ replace: true })).rejects.toThrow('cancelled');
  expect(fs.readFileSync(file(), 'utf8')).toBe(renamed);
  failLogin = false;
  await login({ replace: true });
  expect(read().account.homeAccountId).toBe(bob.homeAccountId);
  expect(fs.readFileSync(file(), 'utf8')).not.toBe(original);
});

test.each([
  { ...alice, environment: 'other-cloud' },
  { ...alice, tenantId: 'other-tenant' },
])('identity binding includes cloud and tenant: %j', async account => {
  await login();
  selected = account;
  await expect(login()).rejects.toThrow('--replace');
});

test('silent refresh is bound to metadata, never falls back, and preserves failed cache changes', async () => {
  await login();
  const original = fs.readFileSync(file(), 'utf8');
  silentAccount = bob;
  await expect(getAccessToken()).rejects.toThrow('refresh failed');
  expect(fs.readFileSync(file(), 'utf8')).toBe(original);
  silentAccount = undefined;
  failSilent = true;
  await expect(getAccessToken()).rejects.toThrow('refresh failed');
  expect(fs.readFileSync(file(), 'utf8')).toBe(original);
  const data = read();
  data.cache = JSON.stringify({ accounts: [bob], revision: 0 });
  fs.writeFileSync(file(), JSON.stringify(data));
  await expect(getAccessToken()).rejects.toThrow('does not match');
  await expect(withProfile('missing', () => getAccessToken())).rejects.toThrow('--profile missing auth login');
});

test('same-profile read/refresh/write transactions serialize without lost updates', async () => {
  await login();
  await Promise.all(Array.from({ length: 8 }, () => getAccessToken()));
  expect(JSON.parse(read().cache).revision).toBe(8);
  expect(fs.existsSync(`${file()}.lock`)).toBe(false);
});

test('logout queues after refresh and cannot be undone by its late write', async () => {
  await login();
  const refresh = getAccessToken();
  await new Promise(resolve => setTimeout(resolve, 1));
  await Promise.all([refresh, logout()]);
  expect(await getAccount()).toBeNull();
});

test('private permissions are repaired and atomic write failures preserve prior state and release locks', async () => {
  await login();
  fs.chmodSync(path.dirname(file()), 0o755);
  fs.chmodSync(file(), 0o644);
  await getAccount();
  expect(fs.statSync(file()).mode & 0o777).toBe(0o600);
  expect(fs.statSync(path.dirname(file())).mode & 0o777).toBe(0o700);
  expect(fs.statSync(path.join(home, '.ms-todo-cli')).mode & 0o777).toBe(0o700);
  const original = fs.readFileSync(file(), 'utf8');
  const rename = fs.renameSync;
  jest.spyOn(fs, 'renameSync').mockImplementation((source, destination) => {
    if (source.toString().endsWith('.new')) throw new Error('Disk failed');
    return rename(source, destination);
  });
  await expect(login()).rejects.toThrow('Disk failed');
  expect(fs.readFileSync(file(), 'utf8')).toBe(original);
  expect(fs.readdirSync(path.dirname(file()))).toEqual(['default.json']);
});

test('a dead process lock is reclaimed, a live process lock is respected', async () => {
  await login();
  fs.mkdirSync(`${file()}.lock`);
  fs.writeFileSync(path.join(`${file()}.lock`, '2147483647-abcdef'), '');
  expect(await getAccessToken()).toBe('test-access-token');
  fs.mkdirSync(`${file()}.lock`);
  const owner = path.join(`${file()}.lock`, `${process.pid}-abcdef`);
  fs.writeFileSync(owner, '');
  const pending = getAccessToken();
  await new Promise(resolve => setTimeout(resolve, 40));
  expect(fs.existsSync(owner)).toBe(true);
  fs.unlinkSync(owner);
  fs.rmdirSync(`${file()}.lock`);
  await expect(pending).resolves.toBe('test-access-token');
});

test('an orphaned empty lock from a crash during acquisition is recovered', async () => {
  await login();
  fs.mkdirSync(`${file()}.lock`);
  const old = new Date(Date.now() - 60000);
  fs.utimesSync(`${file()}.lock`, old, old);
  await expect(getAccessToken()).resolves.toBe('test-access-token');
  expect(fs.existsSync(`${file()}.lock`)).toBe(false);
});

test('account metadata is projected rather than returning arbitrary cached fields', async () => {
  await login();
  const data = JSON.parse(fs.readFileSync(file(), 'utf8'));
  data.account.accessToken = 'not-public';
  data.account.idTokenClaims = { private: 'not-public' };
  fs.writeFileSync(file(), JSON.stringify(data));
  expect(JSON.stringify(await listProfiles())).not.toContain('not-public');
  expect(JSON.stringify(await getAccount())).not.toContain('not-public');
});

test('only unambiguous default legacy caches migrate, without a client ID', async () => {
  const root = path.join(home, '.ms-todo-cli');
  fs.mkdirSync(root);
  const legacy = path.join(root, 'msal-cache.json');
  fs.writeFileSync(legacy, JSON.stringify({ accounts: [alice], revision: 3 }));
  delete process.env.MS_TODO_CLIENT_ID;
  expect(await withProfile('work', () => getAccount())).toBeNull();
  expect(fs.existsSync(legacy)).toBe(true);
  expect(await getAccount()).toMatchObject({ homeAccountId: alice.homeAccountId });
  expect(fs.existsSync(legacy)).toBe(false);
  expect(JSON.parse(read().cache).revision).toBe(3);
  await logout();
  fs.writeFileSync(legacy, JSON.stringify({ accounts: [alice, bob], revision: 0 }));
  await expect(getAccount()).rejects.toThrow('ambiguous');
  expect(fs.existsSync(file())).toBe(false);
  expect(fs.existsSync(legacy)).toBe(true);
  await logout();
  expect(fs.existsSync(legacy)).toBe(false);
});

test('malformed storage is not silently overwritten; unsafe aliases/files are ignored', async () => {
  await login();
  fs.writeFileSync(file(), '{"version":1}');
  await expect(login()).rejects.toThrow('Invalid authentication storage');
  await logout();
  for (const alias of ['..', '.', 'WORK', 'Work', '../escape', 'a/b']) {
    expect(() => withProfile(alias, () => getAccount())).toThrow();
  }
  fs.writeFileSync(path.join(path.dirname(file()), 'WORK.json'), '{}');
  expect(await listProfiles()).toEqual([{ profile: 'default', account: null }]);
});

test('symbolic-link storage is rejected without touching its target', async () => {
  await login();
  const target = path.join(home, 'other.json');
  fs.renameSync(file(), target);
  fs.chmodSync(target, 0o644);
  fs.symlinkSync(target, file());
  await expect(getAccount()).rejects.toThrow('regular file');
  expect(fs.statSync(target).mode & 0o777).toBe(0o644);
});

test('filesystem locks serialize independent processes', async () => {
  await login();
  const program = `
    const Module = require('module');
    const original = Module._load;
    Module._load = function(name, ...args) {
      if (name === '@azure/msal-node') return {PublicClientApplication: class {
        constructor() { this.state = null; }
        getTokenCache() { return {
          deserialize: value => { this.state = JSON.parse(value); },
          serialize: () => JSON.stringify(this.state),
          getAllAccounts: async () => this.state.accounts,
        }; }
        async acquireTokenSilent({account}) {
          const next = this.state.revision + 1;
          await new Promise(resolve => setTimeout(resolve, 40));
          this.state.revision = next;
          return {account, accessToken: 'test-token'};
        }
      }};
      return original.call(this, name, ...args);
    };
    require('./src/auth/authManager').getAccessToken().catch(error => {
      console.error(error); process.exitCode = 1;
    });
  `;
  const run = () => new Promise<void>((resolve, reject) => {
    const child = spawn(process.execPath, ['-r', 'ts-node/register/transpile-only', '-e', program], {
      cwd: process.cwd(), env: { ...process.env, HOME: home }, stdio: ['ignore', 'ignore', 'pipe'],
    });
    let error = '';
    child.stderr.on('data', data => { error += data.toString(); });
    child.on('error', reject);
    child.on('exit', code => code === 0 ? resolve() : reject(new Error(error || `exit ${code}`)));
  });
  await Promise.all([run(), run(), run(), run()]);
  expect(JSON.parse(read().cache).revision).toBe(4);
}, 20000);
