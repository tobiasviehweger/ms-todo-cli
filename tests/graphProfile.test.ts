import { getAccessToken } from '../src/auth/authManager';
import { deleteTask, getLists, getListGroups, getTasks, getTasksAcrossLists, searchTasks, uploadAttachment } from '../src/graph/client';
import { getProfile, withProfile } from '../src/profile';
import { ErrorCodes } from '../src/errors';

jest.mock('../src/auth/authManager', () => ({ getAccessToken: jest.fn() }));

const mockFetch = jest.fn() as jest.MockedFunction<typeof fetch>;
const mockToken = jest.mocked(getAccessToken);
let requests: Array<{ profile: string; url: string; token: string }>;

function jsonResponse(data: unknown, status = 200, headers = new Headers()): Response {
  return {
    ok: status >= 200 && status < 300,
    status,
    statusText: status === 200 ? 'OK' : '',
    headers,
    text: async () => data === undefined ? '' : JSON.stringify(data),
  } as Response;
}

beforeEach(() => {
  jest.clearAllMocks();
  requests = [];
  global.fetch = mockFetch;
  mockToken.mockImplementation(async () => {
    await Promise.resolve();
    return `token-${getProfile()}`;
  });
  mockFetch.mockImplementation(async (input, init) => {
    const url = String(input);
    const headers = init?.headers as Record<string, string>;
    requests.push({
      profile: getProfile(),
      url,
      token: headers.Authorization.split(' ').pop()!,
    });
    if (url.endsWith('createUploadSession')) {
      return jsonResponse({ uploadUrl: 'https://graph.microsoft.com/upload' });
    }
    if (new URL(url).pathname.endsWith('/me/todo/lists')) {
      return jsonResponse({ value: [{ id: 'list', displayName: 'Tasks' }] });
    }
    if (new URL(url).pathname.endsWith('/me/outlook/taskGroups')) {
      return jsonResponse({ value: [] });
    }
    return jsonResponse({ value: [] });
  });
});

test('parallel v1 and beta operations use only their selected accounts', async () => {
  await Promise.all([
    withProfile('work', () => getLists()),
    withProfile('personal', () => getListGroups()),
    withProfile('work', () => searchTasks('review')),
  ]);
  expect(requests.filter((request) => request.url.includes('taskGroups'))[0].token).toBe('token-personal');
  expect(requests.filter((request) => !request.url.includes('taskGroups')).every(
    (request) => request.token === 'token-work',
  )).toBe(true);
});

test('attachment session and each upload request use the selected account', async () => {
  mockFetch.mockImplementation(async (input, init) => {
    const url = String(input);
    if (url.endsWith('createUploadSession')) {
      requests.push({ profile: getProfile(), url, token: '' });
      return jsonResponse({ uploadUrl: 'https://graph.microsoft.com/upload' });
    }
    const headers = init?.headers as Record<string, string>;
    expect(headers.Authorization.split(' ').pop()).toBe('token-work');
    expect(getProfile()).toBe('work');
    requests.push({ profile: getProfile(), url, token: headers.Authorization.split(' ').pop()! });
    return jsonResponse(undefined, 201, new Headers({ location: '/attachments/id' }));
  });
  await withProfile('work', () => uploadAttachment('list', 'task', {
    name: 'note.txt',
    contentBytes: Buffer.from('notes').toString('base64'),
  }));
  expect(requests).toEqual([
    {
      profile: 'work',
      url: 'https://graph.microsoft.com/beta/me/todo/lists/list/tasks/task/attachments/createUploadSession',
      token: '',
    },
    { profile: 'work', url: 'https://graph.microsoft.com/upload/content', token: 'token-work' },
  ]);
  expect(mockFetch.mock.calls[1][1]?.headers).toMatchObject({
    'Content-Type': 'application/octet-stream',
    'Content-Range': 'bytes 0-4/5',
  });
});

test('a Graph client retains its profile when its request starts outside that context', async () => {
  const request = withProfile('work', () => getLists());
  await withProfile('personal', () => request);
  expect(requests[0].profile).toBe('work');
  expect(requests[0].token).toBe('token-work');
});

test('failed authentication never retries under another profile', async () => {
  mockToken.mockImplementation(async () => {
    expect(getProfile()).toBe('missing');
    throw new Error('not authenticated');
  });
  await expect(withProfile('missing', () => getLists())).rejects.toThrow('not authenticated');
  expect(mockToken).toHaveBeenCalledTimes(1);
  expect(requests).toEqual([]);
});

test('maps Graph authentication and missing-resource errors', async () => {
  mockFetch.mockResolvedValueOnce(jsonResponse({ error: { message: 'unauthorized' } }, 401));
  await expect(withProfile('work', () => getLists())).rejects.toMatchObject({
    code: ErrorCodes.AUTH_EXPIRED,
    message: expect.stringContaining('auth login --profile work'),
  });

  mockFetch.mockResolvedValueOnce(jsonResponse({ error: { message: 'not found' } }, 404));
  await expect(withProfile('work', () => getTasks('list', 'Tasks'))).rejects.toMatchObject({
    code: ErrorCodes.LIST_NOT_FOUND,
  });
});

test('handles empty 204 responses for delete requests', async () => {
  mockFetch.mockResolvedValueOnce(jsonResponse(undefined, 204));
  await expect(withProfile('work', () => deleteTask('list', 'task'))).resolves.toBeUndefined();
});

test('pagination and concurrent cross-list requests retain their profile', async () => {
  const pages: Record<string, object> = {
    '/v1.0/me/todo/lists': { value: [{ id: 'one', displayName: 'One' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/lists-page-two' },
    '/v1.0/lists-page-two': { value: [{ id: 'two', displayName: 'Two' }] },
    '/v1.0/me/todo/lists/one/tasks': { value: [], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/tasks-page-two' },
    '/v1.0/tasks-page-two': { value: [] },
    '/v1.0/me/todo/lists/two/tasks': { value: [] },
  };
  mockFetch.mockImplementation(async (input, init) => {
    const url = String(input);
    const headers = init?.headers as Record<string, string>;
    requests.push({
      profile: getProfile(),
      url,
      token: headers.Authorization.split(' ').pop()!,
    });
    return jsonResponse(pages[new URL(url).pathname] || { value: [] });
  });
  await withProfile('work', () => getTasksAcrossLists());
  expect(requests).toHaveLength(5);
  expect(requests.every((request) => request.token === 'token-work')).toBe(true);
});
