import { createProgram, runCli } from '../src/cli';
import { getProfile } from '../src/profile';
import * as auth from '../src/auth';
import * as lists from '../src/commands/list';
import * as groups from '../src/commands/group';
import * as tasks from '../src/commands/task';
import * as steps from '../src/commands/step';
import * as attachments from '../src/commands/attachment';
import { printError } from '../src/output';

jest.mock('../src/auth');
jest.mock('../src/commands/list');
jest.mock('../src/commands/group');
jest.mock('../src/commands/task');
jest.mock('../src/commands/step');
jest.mock('../src/commands/attachment');
jest.mock('../src/output');

const originalProfile = process.env.MS_TODO_PROFILE;
let seen: string[];

beforeEach(() => {
  delete process.env.MS_TODO_PROFILE;
  jest.clearAllMocks();
  seen = [];
  for (const module of [auth, lists, groups, tasks, steps, attachments]) {
    for (const handler of Object.values(module)) {
      if (jest.isMockFunction(handler)) {
        handler.mockImplementation(async () => {
          await Promise.resolve();
          seen.push(getProfile());
        });
      }
    }
  }
});

afterEach(() => {
  if (originalProfile === undefined) delete process.env.MS_TODO_PROFILE;
  else process.env.MS_TODO_PROFILE = originalProfile;
});

const commands = [
  ['auth', 'login'],
  ['auth', 'status'],
  ['auth', 'logout'],
  ['auth', 'print-account'],
  ['auth', 'profiles'],
  ['list', 'create', 'Shopping'],
  ['list', 'list'],
  ['group', 'create', 'Projects'],
  ['group', 'list'],
  ['group', 'update', '--group-id', 'g', '--name', 'Renamed'],
  ['group', 'delete', '--group-id', 'g'],
  ['task', 'create', '--title', 'Title'],
  ['task', 'update', '--task-id', 't', '--title', 'Title'],
  ['task', 'complete', '--task-id', 't'],
  ['task', 'list', '--all-lists'],
  ['task', 'get', '--task-id', 't'],
  ['task', 'attach', '--task-id', 't', '--file', '/tmp/attachment.txt'],
  ['task', 'search', 'keyword'],
  ['step', 'list', '--task-id', 't'],
  ['step', 'create', '--task-id', 't', '--title', 'Title'],
  ['step', 'update', '--task-id', 't', '--step-id', 's', '--title', 'Title'],
  ['step', 'complete', '--task-id', 't', '--step-id', 's'],
  ['step', 'delete', '--task-id', 't', '--step-id', 's'],
  ['create', '--title', 'Title'],
  ['update', '--task-id', 't', '--title', 'Title'],
  ['get', '--task-id', 't'],
];

test.each(commands.map((args) => [args.join(' '), args] as const))(
  '%s accepts a profile before, between, and after subcommands',
  async (_name, args) => {
    const positions = [0, 1, args.length];
    process.env.MS_TODO_PROFILE = 'personal';
    for (const position of positions) {
      const argv = [...args.slice(0, position), '--profile', 'work', ...args.slice(position)];
      await createProgram().parseAsync(argv, { from: 'user' });
    }
    expect(seen).toEqual(['work', 'work', 'work']);
  },
);

test('uses environment or default when the flag is omitted', async () => {
  await createProgram().parseAsync(['auth', 'status'], { from: 'user' });
  process.env.MS_TODO_PROFILE = 'personal';
  await createProgram().parseAsync(['auth', 'status'], { from: 'user' });
  expect(seen).toEqual(['default', 'personal']);
});

test('passes explicit account replacement consent only to login', async () => {
  await createProgram().parseAsync(['auth', 'login', '--profile', 'work', '--replace'], { from: 'user' });
  expect(auth.handleAuthLogin).toHaveBeenCalledWith({ replace: true });
});

test('returns JSON validation errors without running a command for invalid aliases', async () => {
  await runCli(['node', 'cli', 'auth', 'status', '--profile', '../work']);
  expect(printError).toHaveBeenCalledWith('VALIDATION_ERROR', expect.stringContaining('Profile must be'));
  expect(seen).toEqual([]);
});

test('asynchronous action failures are caught by the JSON error handler', async () => {
  jest.mocked(auth.handleAuthStatus).mockRejectedValueOnce(new Error('failed'));
  await runCli(['node', 'cli', 'auth', 'status']);
  expect(printError).toHaveBeenCalledWith('VALIDATION_ERROR', 'Error: failed');
});
