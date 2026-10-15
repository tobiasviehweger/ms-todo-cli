import * as msal from '@azure/msal-node';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { execFileSync } from 'child_process';
import { createHash, randomUUID } from 'crypto';
import { ErrorCodes, AppError } from '../errors';
import { getPinnedAccountIdentity, getProfile, pinAccountIdentity, resolveProfile } from '../profile';

const SCOPES = ['Tasks.ReadWrite', 'User.Read', 'offline_access'];

export interface CachedAccount {
  homeAccountId: string;
  environment: string;
  tenantId: string;
  username: string;
  name: string;
}

interface ProfileData {
  version: 1;
  account: CachedAccount;
  cache: string;
}

function configDir(): string {
  return path.join(os.homedir(), '.ms-todo-cli');
}

function profilesDir(): string {
  return path.join(configDir(), 'profiles');
}

function profileFile(profile: string): string {
  return path.join(profilesDir(), `${resolveProfile(profile)}.json`);
}

function authError(profile: string, message: string): AppError {
  return new AppError(
    ErrorCodes.AUTH_REQUIRED,
    `Profile "${profile}": ${message}. Run: ms-todo-cli --profile ${profile} auth login`,
  );
}

function privateDir(directory: string): void {
  fs.mkdirSync(directory, { recursive: true, mode: 0o700 });
  if (fs.lstatSync(directory).isSymbolicLink()) throw new Error('Authentication directory must not be a symbolic link');
  fs.chmodSync(directory, 0o700);
}

function readPrivate(file: string): string | undefined {
  try {
    if (!fs.lstatSync(file).isFile()) throw new Error('Authentication storage must be a regular file');
    fs.chmodSync(file, 0o600);
    return fs.readFileSync(file, 'utf8');
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === 'ENOENT') return undefined;
    throw err;
  }
}

function atomicWrite(file: string, contents: string): void {
  const temporary = `${file}.${randomUUID()}.new`;
  let descriptor: number | undefined;
  try {
    descriptor = fs.openSync(temporary, 'wx', 0o600);
    fs.writeFileSync(descriptor, contents);
    fs.fsyncSync(descriptor);
    fs.closeSync(descriptor);
    descriptor = undefined;
    fs.renameSync(temporary, file);
  } finally {
    if (descriptor !== undefined) fs.closeSync(descriptor);
    if (fs.existsSync(temporary)) fs.unlinkSync(temporary);
  }
}

function alive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    return (err as NodeJS.ErrnoException).code !== 'ESRCH';
  }
}

function processIncarnation(pid: number): string | undefined {
  try {
    let identity: string | undefined;
    if (process.platform === 'linux') {
      const stat = fs.readFileSync(`/proc/${pid}/stat`, 'utf8');
      const commandEnd = stat.lastIndexOf(')');
      const fields = stat.slice(commandEnd + 2).split(' ');
      const bootId = fs.readFileSync('/proc/sys/kernel/random/boot_id', 'utf8').trim();
      identity = `${bootId}:${fields[19]}`;
    } else if (process.platform === 'darwin') {
      identity = execFileSync('ps', ['-o', 'lstart=', '-p', String(pid)], { encoding: 'utf8' }).trim();
    } else if (process.platform === 'win32') {
      identity = execFileSync('powershell.exe', [
        '-NoProfile',
        '-NonInteractive',
        '-Command',
        `(Get-Process -Id ${pid}).StartTime.ToUniversalTime().Ticks`,
      ], { encoding: 'utf8', timeout: 5000, windowsHide: true }).trim() || undefined;
    }
    return identity ? createHash('sha256').update(identity).digest('hex') : undefined;
  } catch {
    return undefined;
  }
}

async function locked<T>(profile: string, operation: () => Promise<T>): Promise<T> {
  privateDir(configDir());
  privateDir(profilesDir());
  const directory = `${profileFile(profile)}.lock`;
  const owner = `${process.pid}-${processIncarnation(process.pid) || 'unknown'}-${randomUUID()}`;
  const candidate = `${directory}.${owner}.candidate`;
  fs.mkdirSync(candidate, { mode: 0o700 });
  const started = Date.now();
  try {
    fs.writeFileSync(path.join(candidate, owner), '', { flag: 'wx', mode: 0o600 });
    for (;;) {
      try {
        // Publish a nonempty directory atomically, so crash recovery cannot
        // mistake a newly acquired lock for an abandoned empty directory.
        fs.renameSync(candidate, directory);
        break;
      } catch (err) {
        if (!['EEXIST', 'ENOTEMPTY'].includes((err as NodeJS.ErrnoException).code || '')) throw err;
        try {
          if (!fs.lstatSync(directory).isDirectory() || fs.lstatSync(directory).isSymbolicLink()) {
            throw authError(profile, 'Invalid authentication lock');
          }
          const owners = fs.readdirSync(directory);
          if (owners.length === 1) {
            const ownerMatch = /^(\d+)-([a-f0-9]{64}|unknown)-[a-f0-9-]+$/.exec(owners[0]);
            const legacyOwnerMatch = /^(\d+)-[a-f0-9-]+$/.exec(owners[0]);
            const pid = Number(ownerMatch?.[1] || legacyOwnerMatch?.[1]);
            const incarnation = ownerMatch?.[2];
            const currentIncarnation = incarnation && incarnation !== 'unknown' ? processIncarnation(pid) : undefined;
            if ((ownerMatch || legacyOwnerMatch) && (!alive(pid)
              || (currentIncarnation !== undefined && currentIncarnation !== incarnation))) {
              // Only the contender that removes this unique owner may remove the directory.
              fs.unlinkSync(path.join(directory, owners[0]));
              try {
                fs.rmdirSync(directory);
              } catch (cleanupError) {
                if (!['ENOENT', 'ENOTEMPTY'].includes((cleanupError as NodeJS.ErrnoException).code || '')) throw cleanupError;
              }
              continue;
            }
          }
        } catch (cleanupError) {
          if (!['ENOENT', 'ENOTEMPTY'].includes((cleanupError as NodeJS.ErrnoException).code || '')) throw cleanupError;
        }
        if (Date.now() - started > 60000) throw authError(profile, 'Authentication storage is busy; retry shortly');
        await new Promise(resolve => setTimeout(resolve, 25));
      }
    }
    try {
      return await operation();
    } finally {
      fs.unlinkSync(path.join(directory, owner));
      try {
        fs.rmdirSync(directory);
      } catch (err) {
        if (!['ENOENT', 'ENOTEMPTY'].includes((err as NodeJS.ErrnoException).code || '')) throw err;
      }
    }
  } finally {
    if (fs.existsSync(candidate)) fs.rmSync(candidate, { recursive: true });
  }
}

function accountMetadata(account: msal.AccountInfo): CachedAccount {
  if (!account.homeAccountId || !account.environment || !account.tenantId) {
    throw new AppError(ErrorCodes.AUTH_REQUIRED, 'The cached account has no stable identity; sign in again');
  }
  return {
    homeAccountId: account.homeAccountId,
    environment: account.environment,
    tenantId: account.tenantId,
    username: account.username || '',
    name: account.name || '',
  };
}

function sameIdentity(
  a: Pick<CachedAccount, 'homeAccountId' | 'environment' | 'tenantId'>,
  b: Pick<CachedAccount, 'homeAccountId' | 'environment' | 'tenantId'>,
): boolean {
  return a.homeAccountId === b.homeAccountId
    && a.environment.toLowerCase() === b.environment.toLowerCase()
    && a.tenantId === b.tenantId;
}

function readData(profile: string): ProfileData | undefined {
  const contents = readPrivate(profileFile(profile));
  if (contents === undefined) return undefined;
  try {
    const data = JSON.parse(contents) as ProfileData;
    if (data.version !== 1 || typeof data.cache !== 'string' || !data.account
      || !['homeAccountId', 'environment', 'tenantId', 'username', 'name'].every(
        key => typeof (data.account as unknown as Record<string, unknown>)[key] === 'string',
      ) || !data.account.homeAccountId || !data.account.environment || !data.account.tenantId) {
      throw new Error('Invalid profile');
    }
    return { version: 1, cache: data.cache, account: accountMetadata(data.account as msal.AccountInfo) };
  } catch {
    throw authError(profile, 'Invalid authentication storage; log out before signing in again');
  }
}

function createPca(clientId = process.env.MS_TODO_CLIENT_ID): msal.PublicClientApplication {
  if (!clientId) {
    throw new AppError(ErrorCodes.VALIDATION_ERROR, 'MS_TODO_CLIENT_ID environment variable is not set. Register an Azure app and set this variable to its client ID.');
  }
  // Disk writes happen only after the entire operation succeeds and identity is checked.
  return new msal.PublicClientApplication({
    auth: { clientId, authority: 'https://login.microsoftonline.com/common' },
  });
}

async function loadData(profile: string): Promise<ProfileData | undefined> {
  const data = readData(profile);
  if (data || profile !== 'default') return data;
  const legacy = readPrivate(path.join(configDir(), 'msal-cache.json'));
  if (legacy === undefined) return undefined;
  const pca = createPca(process.env.MS_TODO_CLIENT_ID || '00000000-0000-0000-0000-000000000000');
  let migrated: ProfileData;
  try {
    pca.getTokenCache().deserialize(legacy);
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length !== 1) throw new Error('Ambiguous legacy accounts');
    migrated = { version: 1, account: accountMetadata(accounts[0]), cache: legacy };
  } catch {
    throw authError(profile, 'Legacy authentication is ambiguous or invalid; log out and sign in again');
  }
  atomicWrite(profileFile(profile), JSON.stringify(migrated));
  fs.unlinkSync(path.join(configDir(), 'msal-cache.json'));
  return migrated;
}

export async function login(options: { replace?: boolean } = {}): Promise<void> {
  const profile = getProfile();
  await locked(profile, async () => {
    const existing = await loadData(profile);
    // Start with an empty cache: rejected sign-ins must never reach disk.
    const pca = createPca();
    const result = await pca.acquireTokenByDeviceCode({
      scopes: SCOPES,
      deviceCodeCallback: response => process.stderr.write(`[${profile}] ${response.message}\n`),
    });
    if (!result?.account || !result.accessToken) throw authError(profile, 'Login failed: no account or token returned');
    const account = accountMetadata(result.account);
    if (existing && !sameIdentity(existing.account, account) && !options.replace) {
      throw authError(profile, 'A different account is already bound; use auth login --replace to replace it');
    }
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length !== 1 || !sameIdentity(accountMetadata(accounts[0]), account)) {
      throw authError(profile, 'Login returned an ambiguous account cache');
    }
    atomicWrite(profileFile(profile), JSON.stringify({ version: 1, account, cache: pca.getTokenCache().serialize() }));
  });
}

export async function getAccessToken(): Promise<string> {
  const profile = getProfile();
  return locked(profile, async () => {
    const data = await loadData(profile);
    if (!data) throw authError(profile, 'Not authenticated');
    const pinnedIdentity = getPinnedAccountIdentity();
    if (pinnedIdentity && !sameIdentity(pinnedIdentity, data.account)) {
      throw new AppError(ErrorCodes.AUTH_REQUIRED, `Profile "${profile}": account binding changed during this command; retry the command`);
    }
    pinAccountIdentity(data.account);
    const pca = createPca();
    let account: msal.AccountInfo;
    try {
      pca.getTokenCache().deserialize(data.cache);
      const accounts = await pca.getTokenCache().getAllAccounts();
      const matches = accounts.filter(cached => sameIdentity(accountMetadata(cached), data.account));
      if (matches.length !== 1 || accounts.length !== 1) throw authError(profile, 'Cached account does not match the profile identity');
      account = matches[0];
    } catch (err) {
      if (err instanceof AppError) throw err;
      throw authError(profile, 'Cached authentication is invalid; log out and sign in again');
    }
    let result: msal.AuthenticationResult | null;
    try {
      result = await pca.acquireTokenSilent({ account, scopes: SCOPES });
      if (!result?.accessToken || !result.account || !sameIdentity(accountMetadata(result.account), data.account)) {
        throw new Error('Token identity mismatch');
      }
    } catch {
      throw new AppError(ErrorCodes.AUTH_EXPIRED, `Profile "${profile}": token refresh failed. Run: ms-todo-cli --profile ${profile} auth login`);
    }
    atomicWrite(profileFile(profile), JSON.stringify({ ...data, cache: pca.getTokenCache().serialize() }));
    return result.accessToken;
  });
}

export async function logout(): Promise<void> {
  const profile = getProfile();
  await locked(profile, async () => {
    const files = [profileFile(profile)];
    if (profile === 'default') files.push(path.join(configDir(), 'msal-cache.json'));
    for (const file of files) {
      try {
        fs.unlinkSync(file);
      } catch (err) {
        if ((err as NodeJS.ErrnoException).code !== 'ENOENT') throw err;
      }
    }
  });
}

export async function getAccount(): Promise<CachedAccount | null> {
  const profile = getProfile();
  return locked(profile, async () => (await loadData(profile))?.account || null);
}

export async function isAuthenticated(): Promise<boolean> {
  return (await getAccount()) !== null;
}

export async function listProfiles(): Promise<{ profile: string; account: CachedAccount | null }[]> {
  let files: string[];
  try {
    files = fs.readdirSync(profilesDir());
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code !== 'ENOENT') throw err;
    files = [];
  }
  const aliases = new Set<string>(['default', getProfile()]);
  for (const file of files) {
    if (!file.endsWith('.json')) continue;
    const alias = file.slice(0, -5);
    try {
      if (resolveProfile(alias) === alias) aliases.add(alias);
    } catch {
      // Untrusted filenames are never used as profile aliases.
    }
  }
  const profiles: { profile: string; account: CachedAccount | null }[] = [];
  for (const profile of [...aliases].sort()) {
    const account = await locked(profile, async () => (await loadData(profile))?.account || null);
    if (account || profile === getProfile()) profiles.push({ profile, account });
  }
  return profiles;
}
