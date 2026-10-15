import axios, { AxiosInstance, InternalAxiosRequestConfig } from 'axios';
import { getAccessToken } from '../src/auth/authManager';
import { getLists, getListGroups, getTasksAcrossLists, searchTasks, uploadAttachment } from '../src/graph/client';
import { getProfile, withProfile } from '../src/profile';
import { AppError, ErrorCodes } from '../src/errors';

jest.mock('axios');
jest.mock('../src/auth/authManager', () => ({ getAccessToken: jest.fn() }));

const mockAxios = jest.mocked(axios);
const mockToken = jest.mocked(getAccessToken);
let requests: Array<{ profile: string; url: string }>;

beforeEach(() => {
  jest.clearAllMocks();
  requests = [];
  mockToken.mockImplementation(async () => `token-${getProfile()}`);
  mockAxios.create.mockImplementation(() => {
    let interceptor: (config: InternalAxiosRequestConfig) => Promise<InternalAxiosRequestConfig>;
    const request = async (url: string) => {
      const config = await interceptor({ headers: {} } as InternalAxiosRequestConfig);
      requests.push({ profile: String(config.headers.Authorization).split(' ').pop()!, url });
      if (url.endsWith('createUploadSession')) {
        return { data: { uploadUrl: 'https://graph.microsoft.com/upload' } };
      }
      if (url === '/me/todo/lists') {
        return { data: { value: [{ id: 'list', displayName: 'Tasks' }] } };
      }
      return { data: { value: [] } };
    };
    return {
      get: jest.fn(request),
      post: jest.fn(request),
      interceptors: {
        request: { use: jest.fn((handler) => { interceptor = handler; }) },
        response: { use: jest.fn() },
      },
    } as unknown as AxiosInstance;
  });
});

test('parallel v1 and beta operations use only their selected accounts', async () => {
  await Promise.all([
    withProfile('work', () => getLists()),
    withProfile('personal', () => getListGroups()),
    withProfile('work', () => searchTasks('review')),
  ]);
  expect(requests.filter((r) => r.url.includes('taskGroups'))[0].profile).toBe('token-personal');
  expect(requests.filter((r) => !r.url.includes('taskGroups')).every((r) => r.profile === 'token-work')).toBe(true);
});

test('attachment session and each upload request use the selected account', async () => {
  mockAxios.put.mockImplementation(async (_url, _body, config) => {
    expect(String(config?.headers?.Authorization).split(' ').pop()).toBe('token-work');
    expect(getProfile()).toBe('work');
    return { status: 201, headers: { location: '/attachments/id' }, data: {} };
  });
  await withProfile('work', () => uploadAttachment('list', 'task', {
    name: 'note.txt',
    contentBytes: Buffer.from('notes').toString('base64'),
  }));
  expect(requests).toEqual([{
    profile: 'token-work',
    url: '/me/todo/lists/list/tasks/task/attachments/createUploadSession',
  }]);
  expect(mockAxios.put).toHaveBeenCalledTimes(1);
});

test('an interceptor retains its profile even if called outside the originating context', async () => {
  await withProfile('work', () => getLists());
  const client = mockAxios.create.mock.results[0].value as AxiosInstance;
  const use = client.interceptors.request.use as jest.Mock;
  const interceptor = use.mock.calls[0][0];
  const config = await withProfile('personal', () => interceptor({ headers: {} }));
  expect(String(config.headers.Authorization).split(' ').pop()).toBe('token-work');
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

test('Graph response interceptors preserve profile-specific authentication errors', async () => {
  await withProfile('work', () => getLists());
  const client = mockAxios.create.mock.results[0].value as AxiosInstance;
  const interceptor = (client.interceptors.response.use as jest.Mock).mock.calls[0][1];
  const error = new AppError(ErrorCodes.AUTH_REQUIRED, 'Profile work requires login');
  expect(() => interceptor(error)).toThrow(error);
  expect(() => interceptor({ response: { status: 401 } })).toThrow('auth login --profile work');
});

test('pagination and concurrent cross-list requests retain their profile', async () => {
  const pages: Record<string, object> = {
    '/me/todo/lists': { value: [{ id: 'one', displayName: 'One' }], '@odata.nextLink': '/lists-page-two' },
    '/lists-page-two': { value: [{ id: 'two', displayName: 'Two' }] },
    '/me/todo/lists/one/tasks': { value: [], '@odata.nextLink': '/tasks-page-two' },
    '/tasks-page-two': { value: [] },
    '/me/todo/lists/two/tasks': { value: [] },
  };
  mockAxios.create.mockImplementationOnce(() => {
    let interceptor: (config: InternalAxiosRequestConfig) => Promise<InternalAxiosRequestConfig>;
    return {
      get: jest.fn(async (url: string) => {
        await Promise.resolve();
        const config = await interceptor({ headers: {} } as InternalAxiosRequestConfig);
        requests.push({ profile: String(config.headers.Authorization).split(' ').pop()!, url });
        return { data: pages[url] };
      }),
      interceptors: {
        request: { use: jest.fn((handler) => { interceptor = handler; }) },
        response: { use: jest.fn() },
      },
    } as unknown as AxiosInstance;
  });
  await withProfile('work', () => getTasksAcrossLists());
  expect(requests).toHaveLength(5);
  expect(requests.every((request) => request.profile === 'token-work')).toBe(true);
});
