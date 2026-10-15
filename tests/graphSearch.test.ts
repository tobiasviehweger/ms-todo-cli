import { createListGroup, getListGroups, getLists, searchTasks, updateListGroup } from '../src/graph/client';

jest.mock('../src/auth/authManager', () => ({
  getAccessToken: jest.fn().mockResolvedValue('token'),
}));

const mockFetch = jest.fn() as jest.MockedFunction<typeof fetch>;

function jsonResponse(data: unknown): Response {
  return {
    ok: true,
    status: 200,
    statusText: 'OK',
    headers: new Headers({ 'content-type': 'application/json' }),
    text: async () => JSON.stringify(data),
  } as Response;
}

beforeEach(() => {
  jest.clearAllMocks();
  global.fetch = mockFetch;
});

describe('searchTasks', () => {
  test('normalizes search terms before building filter params', async () => {
    mockFetch.mockResolvedValueOnce(jsonResponse({ value: [{ id: 'list-1', displayName: 'Inbox' }] }));
    mockFetch.mockResolvedValueOnce(jsonResponse({ value: [] }));

    await searchTasks('  FoO!!!   bar  ');

    const url = new URL(String(mockFetch.mock.calls[1][0]));
    expect(url.searchParams.get('@term')).toBe("'foo bar'");
    expect(url.searchParams.get('$filter')).toContain('@term');
  });

  test('handles search terms that normalize to an empty string', async () => {
    await expect(searchTasks('  !!!   ???  ')).resolves.toEqual([]);
    expect(mockFetch).not.toHaveBeenCalled();
  });

  test('handles whitespace-only search terms', async () => {
    await expect(searchTasks('     ')).resolves.toEqual([]);
    expect(mockFetch).not.toHaveBeenCalled();
  });

  test('preserves and normalizes unicode search terms consistently', async () => {
    mockFetch.mockResolvedValueOnce(jsonResponse({ value: [{ id: 'list-1', displayName: 'Inbox' }] }));
    mockFetch.mockResolvedValueOnce(jsonResponse({ value: [] }));

    await searchTasks('  Café   résumé  ');

    expect(new URL(String(mockFetch.mock.calls[1][0])).searchParams.get('@term')).toBe("'café résumé'");
  });
});

describe('Graph requests', () => {
  test('uses fetch to get lists from the v1.0 Graph endpoint', async () => {
    mockFetch.mockResolvedValueOnce(jsonResponse({ value: [{ id: 'list-1', displayName: 'Inbox' }] }));

    await expect(getLists()).resolves.toEqual([{ id: 'list-1', displayName: 'Inbox' }]);
    expect(String(mockFetch.mock.calls[0][0])).toBe('https://graph.microsoft.com/v1.0/me/todo/lists');
    const headers = mockFetch.mock.calls[0][1]?.headers as Record<string, string>;
    expect(headers['Content-Type']).toBe('application/json');
  });

  test('maps beta task group fields and serializes JSON mutation bodies', async () => {
    mockFetch.mockResolvedValueOnce(jsonResponse({
      value: [{ id: 'group-1', name: 'Work', isDefaultGroup: false, groupKey: 'k1', changeKey: 'c1' }],
    }));
    await expect(getListGroups()).resolves.toEqual([{ id: 'group-1', displayName: 'Work' }]);
    expect(String(mockFetch.mock.calls[0][0])).toBe('https://graph.microsoft.com/beta/me/outlook/taskGroups');

    mockFetch.mockResolvedValueOnce(jsonResponse({
      id: 'group-2', name: 'Personal', isDefaultGroup: false, groupKey: 'k2', changeKey: 'c2',
    }));
    await expect(createListGroup('Personal')).resolves.toEqual({ id: 'group-2', displayName: 'Personal' });
    expect(mockFetch.mock.calls[1][1]?.method).toBe('POST');
    expect(mockFetch.mock.calls[1][1]?.body).toBe(JSON.stringify({ name: 'Personal' }));

    mockFetch.mockResolvedValueOnce(jsonResponse({
      id: 'group-3', name: 'Renamed', isDefaultGroup: false, groupKey: 'k3', changeKey: 'c3',
    }));
    await expect(updateListGroup('group-3', { displayName: 'Renamed' })).resolves.toEqual({
      id: 'group-3',
      displayName: 'Renamed',
    });
    expect(mockFetch.mock.calls[2][1]?.method).toBe('PATCH');
    expect(mockFetch.mock.calls[2][1]?.body).toBe(JSON.stringify({ name: 'Renamed' }));
  });
});
