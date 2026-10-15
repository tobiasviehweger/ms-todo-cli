import { EventEmitter } from 'events';
import * as readline from 'readline';
import * as graph from '../src/graph/client';
import { handleTaskCreate, handleTaskUpdate } from '../src/commands/task';
import { getProfile, withProfile } from '../src/profile';
import { printSuccess } from '../src/output';

jest.mock('readline', () => ({ createInterface: jest.fn() }));
jest.mock('../src/graph/client');
jest.mock('../src/output');

beforeEach(() => {
  jest.clearAllMocks();
  jest.mocked(readline.createInterface).mockImplementation(() => {
    const input = new EventEmitter();
    queueMicrotask(() => {
      input.emit('line', JSON.stringify({
        profile: 'personal', title: 'Title', taskId: 'task', list: 'Inbox',
      }));
      input.emit('close');
    });
    return input as readline.Interface;
  });
});

test('stdin task creation cannot switch profiles during list resolution or creation', async () => {
  jest.mocked(graph.getListByName).mockImplementation(async () => {
    expect(getProfile()).toBe('work');
    return { id: 'list', displayName: 'Inbox' };
  });
  const task = { id: 'task', title: 'Title' };
  jest.mocked(graph.createTask).mockImplementation(async () => {
    expect(getProfile()).toBe('work');
    return task;
  });
  await withProfile('work', () => handleTaskCreate({ stdin: true }));
  expect(graph.createTask).toHaveBeenCalled();
  expect(printSuccess).toHaveBeenCalledWith({ task });
});

test('stdin task updates cannot switch profiles during lookup or mutation', async () => {
  const task = { id: 'task', title: 'Title' };
  jest.mocked(graph.findTaskById).mockImplementation(async () => {
    expect(getProfile()).toBe('work');
    return { task, listId: 'list' };
  });
  jest.mocked(graph.updateTask).mockImplementation(async () => {
    expect(getProfile()).toBe('work');
    return task;
  });
  await withProfile('work', () => handleTaskUpdate({ stdin: true }));
  expect(graph.updateTask).toHaveBeenCalled();
  expect(printSuccess).toHaveBeenCalledWith({ task });
});
