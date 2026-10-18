import { runCli } from '../src/cli';
import * as output from '../src/output';

let stdout: jest.SpyInstance;
let stderr: jest.SpyInstance;
let exit: jest.SpyInstance;
let printError: jest.SpyInstance;

beforeEach(() => {
  stdout = jest.spyOn(process.stdout, 'write').mockImplementation(() => true);
  stderr = jest.spyOn(process.stderr, 'write').mockImplementation(() => true);
  exit = jest.spyOn(process, 'exit').mockImplementation(() => undefined as never);
  printError = jest.spyOn(output, 'printError').mockImplementation(() => {});
});

afterEach(() => {
  jest.restoreAllMocks();
});

function text(stream: jest.SpyInstance): string {
  return stream.mock.calls.map(([chunk]) => String(chunk)).join('');
}

test.each(['--help', '-h', 'help'])('%s shows top-level commands', async (flag) => {
  await runCli(['node', 'cli', flag]);

  expect(text(stdout)).toContain('Usage: ms-todo-cli');
  expect(text(stdout)).toContain('task');
  expect(text(stdout)).toContain('--profile <alias>');
  expect(text(stderr)).toBe('');
  expect(exit).toHaveBeenCalledWith(0);
  expect(printError).not.toHaveBeenCalled();
});

test.each(['auth', 'list', 'group', 'task', 'step'])(
  '%s --help shows subcommands',
  async (group) => {
    await runCli(['node', 'cli', group, '--help']);

    expect(text(stdout)).toContain(`Usage: ms-todo-cli ${group}`);
    expect(text(stdout)).toContain('Commands:');
    expect(text(stdout)).toContain('--profile <alias>');
    expect(text(stderr)).toBe('');
    expect(exit).toHaveBeenCalledWith(0);
    expect(printError).not.toHaveBeenCalled();
  },
);

test.each([
  ['task', 'create', '--help'],
  ['task', 'help', 'create'],
  ['create', '-h'],
])('%s %s %s shows task parameters', async (...args) => {
  await runCli(['node', 'cli', ...args]);

  for (const option of ['--title <title>', '--list <name>', '--due <date>', '--stdin', '--profile <alias>']) {
    expect(text(stdout)).toContain(option);
  }
  expect(text(stderr)).toBe('');
  expect(exit).toHaveBeenCalledWith(0);
  expect(printError).not.toHaveBeenCalled();
});

test('help works without required command options or a valid profile', async () => {
  await runCli(['node', 'cli', 'task', 'attach', '--profile', '../invalid', '--help']);

  expect(text(stdout)).toContain('--task-id <id>');
  expect(text(stdout)).toContain('--file <path>');
  expect(exit).toHaveBeenCalledWith(0);
  expect(printError).not.toHaveBeenCalled();
});

test.each([[], ['auth'], ['list'], ['group'], ['task'], ['step']])(
  'missing subcommands show usage instead of a JSON error (%j)',
  async (...args) => {
    await runCli(['node', 'cli', ...args]);

    expect(text(stderr)).toContain('Usage: ms-todo-cli');
    expect(text(stderr)).toContain('Commands:');
    expect(text(stdout)).toBe('');
    expect(exit).toHaveBeenCalledWith(1);
    expect(printError).not.toHaveBeenCalled();
  },
);

test.each(['--version', '-V'])('%s prints the version', async (flag) => {
  await runCli(['node', 'cli', flag]);

  expect(text(stdout)).toBe('0.1.0\n');
  expect(text(stderr)).toBe('');
  expect(exit).toHaveBeenCalledWith(0);
  expect(printError).not.toHaveBeenCalled();
});

test.each([
  ['unknown'],
  ['task', 'create', '--unknown'],
  ['task', 'complete'],
  ['task', 'create', '--title'],
  ['list', 'create'],
])('invalid arguments still produce only JSON errors (%j)', async (...args) => {
  await runCli(['node', 'cli', ...args]);

  expect(printError).toHaveBeenCalledWith('VALIDATION_ERROR', expect.any(String));
  expect(text(stdout)).toBe('');
  expect(text(stderr)).toBe('');
});
