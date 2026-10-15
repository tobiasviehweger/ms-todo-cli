import { getAccessToken as acquireAccessToken } from '../auth/authManager';
import { getProfile, withProfile } from '../profile';
import { ErrorCodes, AppError } from '../errors';
import { sanitizeSearchTerm } from '../utils/search';
import { TodoTask, TodoList, ChecklistItem, TaskAttachment, TodoListGroup } from '../schema/types';

const BASE_URL = 'https://graph.microsoft.com/v1.0';
const BETA_BASE_URL = 'https://graph.microsoft.com/beta';
// Small batch size balances latency while reducing Graph API throttling risk.
const TASK_FETCH_BATCH_SIZE = 3;

interface GraphResponse<T> {
  data: T;
  status: number;
  headers: Headers;
}

interface GraphClient {
  get<T = unknown>(url: string, config?: { params?: Record<string, string> }): Promise<GraphResponse<T>>;
  post<T = unknown>(url: string, data?: unknown): Promise<GraphResponse<T>>;
  patch<T = unknown>(url: string, data?: unknown): Promise<GraphResponse<T>>;
  delete<T = unknown>(url: string): Promise<GraphResponse<T>>;
}

function createGraphClient(baseURL: string): GraphClient {
  const profile = getProfile();
  const request = async <T>(
    method: string,
    rawUrl: string,
    body?: unknown,
    params?: Record<string, string>,
  ): Promise<GraphResponse<T>> => {
    const token = await withProfile(profile, () => acquireAccessToken());
    const url = /^https?:\/\//i.test(rawUrl)
      ? new URL(rawUrl)
      : new URL(`${baseURL.replace(/\/+$/, '')}/${rawUrl.replace(/^\/+/, '')}`);
    if (params) {
      for (const [key, value] of Object.entries(params)) url.searchParams.set(key, value);
    }
    let response: Response;
    try {
      response = await fetch(url, {
        method,
        headers: {
          Authorization: ['Bearer', token].join(' '),
          'Content-Type': 'application/json',
        },
        body: body === undefined ? undefined : JSON.stringify(body),
      });
    } catch (err) {
      throw new AppError(ErrorCodes.GRAPH_ERROR, err instanceof Error ? err.message : String(err));
    }
    const text = response.status === 204 ? '' : await response.text();
    let data: unknown;
    if (text) {
      try {
        data = JSON.parse(text);
      } catch {
        data = text;
      }
    }
    if (!response.ok) {
      if (response.status === 429) throw new AppError(ErrorCodes.RATE_LIMITED, 'Rate limited by Microsoft Graph API');
      if (response.status === 401 || response.status === 403) {
        throw new AppError(ErrorCodes.AUTH_EXPIRED, `Authentication expired or invalid for profile "${profile}". Run: ms-todo-cli auth login --profile ${profile}`);
      }
      if (response.status === 404) throw mapNotFoundError(normalizeGraphUrl(rawUrl, baseURL));
      const error = data && typeof data === 'object' && 'error' in data ? data.error : undefined;
      const message = error && typeof error === 'object' && 'message' in error ? error.message : undefined;
      throw new AppError(
        ErrorCodes.GRAPH_ERROR,
        typeof message === 'string' ? message : response.statusText || `HTTP ${response.status}`,
      );
    }
    return { data: data as T, status: response.status, headers: response.headers };
  };
  const client: GraphClient = {
    get: <T>(url: string, config?: { params?: Record<string, string> }) => request<T>('GET', url, undefined, config?.params),
    post: <T>(url: string, data?: unknown) => request<T>('POST', url, data),
    patch: <T>(url: string, data?: unknown) => request<T>('PATCH', url, data),
    delete: <T>(url: string) => request<T>('DELETE', url),
  };
  return client;
}

function normalizeGraphUrl(rawUrl: string, baseUrl: string = BASE_URL): string {
  if (!rawUrl) return '';
  try {
    const parsed = new URL(rawUrl, baseUrl);
    return parsed.pathname;
  } catch {
    const withoutFragment = rawUrl.split('#')[0];
    return withoutFragment.split('?')[0];
  }
}

function mapNotFoundError(url: string): AppError {
  if (/\/taskGroups\/[^/]+/.test(url)) {
    return new AppError(ErrorCodes.LIST_GROUP_NOT_FOUND, 'List group not found');
  }
  if (/\/lists\/[^/]+(?:\/tasks)?$/.test(url)) {
    return new AppError(ErrorCodes.LIST_NOT_FOUND, 'List not found');
  }
  return new AppError(ErrorCodes.TASK_NOT_FOUND, 'Resource not found');
}

function createBetaClient(): GraphClient {
  return createGraphClient(BETA_BASE_URL);
}

function createClient(): GraphClient {
  return createGraphClient(BASE_URL);
}

type GraphTask = {
  id: string;
  title: string;
  status?: string;
  body?: { content?: string };
  dueDateTime?: { dateTime?: string };
  importance?: string;
  completedDateTime?: { dateTime?: string };
};

const MAX_CONCURRENT_SEARCH_REQUESTS = 5;
const ODATA_TERM_PARAMETER = '@term';

function mapTask(item: GraphTask, listName?: string, listId?: string): TodoTask {
  return {
    id: item.id,
    title: item.title,
    list: listName,
    listId: listId,
    status: item.status,
    notes: item.body?.content,
    dueDateTime: item.dueDateTime?.dateTime,
    priority: item.importance,
    completedDateTime: item.completedDateTime?.dateTime,
  };
}

type ODataPage<T> = { value?: T[]; '@odata.nextLink'?: string };
type PageItemMapper<TInput, TOutput> = (item: TInput) => TOutput;

async function mapConcurrent<TInput, TOutput>(
  items: TInput[],
  concurrency: number,
  worker: (item: TInput) => Promise<TOutput>,
): Promise<TOutput[]> {
  if (items.length === 0) {
    return [];
  }

  const results: TOutput[] = new Array(items.length);
  const workQueue: Array<{ index: number; item: TInput }> = items.map((item, index) => ({ index, item }));

  async function runWorker(): Promise<void> {
    while (workQueue.length > 0) {
      const nextWorkItem = workQueue.pop();
      if (!nextWorkItem) {
        return;
      }

      results[nextWorkItem.index] = await worker(nextWorkItem.item);
    }
  }

  const workerCount = Math.min(Math.max(1, concurrency), items.length);
  await Promise.all(Array.from({ length: workerCount }, () => runWorker()));
  return results;
}

async function fetchPaged<TInput, TOutput>(
  client: GraphClient,
  url: string,
  mapItem: PageItemMapper<TInput, TOutput>,
): Promise<TOutput[]> {
  const results: TOutput[] = [];
  let nextUrl: string | undefined = url;
  while (nextUrl) {
    const res: GraphResponse<ODataPage<TInput>> = await client.get<ODataPage<TInput>>(nextUrl);
    for (const item of res.data.value || []) {
      results.push(mapItem(item));
    }
    nextUrl = res.data['@odata.nextLink'];
  }
  return results;
}

function buildTaskSearchFilter(): string {
  return `contains(tolower(title),${ODATA_TERM_PARAMETER}) or contains(tolower(body/content),${ODATA_TERM_PARAMETER})`;
}

function buildTaskSearchParams(term: string): Record<string, string> {
  const escaped = term.replace(/'/g, "''");
  return {
    $filter: buildTaskSearchFilter(),
    $select: 'id,title,status,body,dueDateTime,importance,completedDateTime',
    [ODATA_TERM_PARAMETER]: `'${escaped}'`,
  };
}

async function fetchSearchTasksForList(
  client: GraphClient,
  list: TodoList,
  params: Record<string, string>,
): Promise<TodoTask[]> {
  const tasks: TodoTask[] = [];
  let nextUrl: string | undefined = `/me/todo/lists/${list.id}/tasks`;
  let isFirstRequest = true;

  while (nextUrl) {
    const res: GraphResponse<ODataPage<GraphTask>> = await client.get<ODataPage<GraphTask>>(
      nextUrl,
      isFirstRequest ? { params } : undefined,
    );
    const items = res.data.value || [];
    tasks.push(...items.map((item) => mapTask(item as GraphTask, list.displayName, list.id)));
    nextUrl = res.data['@odata.nextLink'];
    isFirstRequest = false;
  }

  return tasks;
}

interface OutlookTaskGroup {
  id: string;
  name: string;
  isDefaultGroup: boolean;
  groupKey: string;
  changeKey: string;
}

function mapOutlookGroup(g: OutlookTaskGroup): TodoListGroup {
  return { id: g.id, displayName: g.name };
}

export async function getListGroups(): Promise<TodoListGroup[]> {
  const client = createBetaClient();
  return fetchPaged<OutlookTaskGroup, TodoListGroup>(client, '/me/outlook/taskGroups', mapOutlookGroup);
}

export async function getListGroupByName(name: string): Promise<TodoListGroup | null> {
  const groups = await getListGroups();
  return groups.find((g) => g.displayName.toLowerCase() === name.toLowerCase()) || null;
}

export async function createListGroup(displayName: string): Promise<TodoListGroup> {
  const client = createBetaClient();
  const res = await client.post<OutlookTaskGroup>('/me/outlook/taskGroups', { name: displayName });
  return mapOutlookGroup(res.data);
}

export async function updateListGroup(listGroupId: string, updates: { displayName?: string }): Promise<TodoListGroup> {
  if (!updates.displayName?.trim()) {
    throw new AppError(ErrorCodes.VALIDATION_ERROR, 'displayName is required');
  }
  const client = createBetaClient();
  const res = await client.patch<OutlookTaskGroup>(`/me/outlook/taskGroups/${listGroupId}`, { name: updates.displayName.trim() });
  return mapOutlookGroup(res.data);
}

export async function deleteListGroup(listGroupId: string): Promise<void> {
  const client = createBetaClient();
  await client.delete(`/me/outlook/taskGroups/${listGroupId}`);
}

export async function getLists(client?: GraphClient): Promise<TodoList[]> {
  const activeClient = client ?? createClient();
  const res = await activeClient.get<{ value: TodoList[] }>('/me/todo/lists');
  return res.data.value;
}

export async function getListByName(name: string): Promise<TodoList | null> {
  const lists = await getLists();
  return lists.find((l) => l.displayName.toLowerCase() === name.toLowerCase()) || null;
}

export async function createList(displayName: string): Promise<TodoList> {
  const client = createClient();
  const res = await client.post<TodoList>('/me/todo/lists', { displayName });
  return res.data;
}

export async function getTasks(listId: string, listName?: string): Promise<TodoTask[]> {
  const client = createClient();
  if (listName === undefined) {
    const lists = await getLists();
    const list = lists.find((l) => l.id === listId);
    listName = list?.displayName;
  }
  const res = await client.get<{ value?: GraphTask[] }>(`/me/todo/lists/${listId}/tasks`);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return (res.data.value || []).map((t: any) => mapTask(t, listName, listId));
}

export async function searchTasks(keyword: string): Promise<TodoTask[]> {
  const client = createClient();
  const normalized = sanitizeSearchTerm(keyword);
  if (!normalized) {
    return [];
  }
  const lists = await getLists(client);
  const term = normalized.toLowerCase();
  const params = buildTaskSearchParams(term);

  const results = await mapConcurrent(lists, MAX_CONCURRENT_SEARCH_REQUESTS, (list) => {
    return fetchSearchTasksForList(client, list, params);
  });

  return results.flat();
}

/** Internal helper: fetch a single task without calling getLists(). */
async function fetchTaskRaw(
  client: GraphClient,
  listId: string,
  taskId: string,
  listName?: string,
): Promise<TodoTask> {
  const res = await client.get<GraphTask>(`/me/todo/lists/${listId}/tasks/${taskId}`);
  return mapTask(res.data, listName, listId);
}

export async function getTask(listId: string, taskId: string, listName?: string): Promise<TodoTask> {
  const client = createClient();
  if (listName === undefined) {
    const lists = await getLists();
    const list = lists.find((l) => l.id === listId);
    listName = list?.displayName;
  }
  return fetchTaskRaw(client, listId, taskId, listName);
}

export async function createTask(listId: string, task: object, listName?: string): Promise<TodoTask> {
  const client = createClient();
  if (listName === undefined) {
    const lists = await getLists();
    const list = lists.find((l) => l.id === listId);
    listName = list?.displayName;
  }
  const res = await client.post<GraphTask>(`/me/todo/lists/${listId}/tasks`, task);
  return mapTask(res.data, listName, listId);
}

export async function updateTask(listId: string, taskId: string, updates: object, listName?: string): Promise<TodoTask> {
  const client = createClient();
  if (listName === undefined) {
    const lists = await getLists();
    const list = lists.find((l) => l.id === listId);
    listName = list?.displayName;
  }
  const res = await client.patch<GraphTask>(`/me/todo/lists/${listId}/tasks/${taskId}`, updates);
  return mapTask(res.data, listName, listId);
}

export async function deleteTask(listId: string, taskId: string): Promise<void> {
  const client = createClient();
  await client.delete(`/me/todo/lists/${listId}/tasks/${taskId}`);
}

/**
 * Find a task by its ID across all lists.
 * If listId is provided, only that list is checked (O(1) network calls).
 * Otherwise, getLists() is called ONCE and each list is probed with a single client.
 */
export async function findTaskById(taskId: string, listId?: string): Promise<{ task: TodoTask; listId: string } | null> {
  const client = createClient();
  if (listId) {
    try {
      const task = await fetchTaskRaw(client, listId, taskId);
      return { task, listId };
    } catch (err) {
      if (err instanceof AppError && err.code === ErrorCodes.TASK_NOT_FOUND) return null;
      throw err;
    }
  }
  const lists = await getLists();
  for (const list of lists) {
    try {
      const task = await fetchTaskRaw(client, list.id, taskId, list.displayName);
      return { task, listId: list.id };
    } catch (err) {
      if (err instanceof AppError && err.code === ErrorCodes.TASK_NOT_FOUND) continue;
      throw err;
    }
  }
  return null;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function mapChecklistItem(item: any): ChecklistItem {
  return {
    id: item.id,
    displayName: item.displayName,
    isChecked: item.isChecked ?? false,
    checkedDateTime: item.checkedDateTime,
  };
}

export async function getChecklistItems(listId: string, taskId: string): Promise<ChecklistItem[]> {
  const client = createClient();
  const res = await client.get<{ value?: ChecklistItem[] }>(`/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`);
  return (res.data.value || []).map(mapChecklistItem);
}

export async function createChecklistItem(listId: string, taskId: string, displayName: string): Promise<ChecklistItem> {
  const client = createClient();
  const res = await client.post<ChecklistItem>(`/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`, { displayName });
  return mapChecklistItem(res.data);
}

export async function updateChecklistItem(
  listId: string,
  taskId: string,
  checklistItemId: string,
  updates: { displayName?: string; isChecked?: boolean },
): Promise<ChecklistItem> {
  const client = createClient();
  const res = await client.patch<ChecklistItem>(
    `/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
    updates,
  );
  return mapChecklistItem(res.data);
}

export async function deleteChecklistItem(listId: string, taskId: string, checklistItemId: string): Promise<void> {
  const client = createClient();
  await client.delete(`/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`);
}

export async function uploadAttachment(
  listId: string,
  taskId: string,
  attachment: { name: string; contentBytes: string; contentType?: string },
): Promise<TaskAttachment> {
  const profile = getProfile();
  const getAccessToken = () => withProfile(profile, () => acquireAccessToken());
  const betaClient = createBetaClient();
  const contentBytes = Buffer.from(attachment.contentBytes, 'base64');
  const fileSize = contentBytes.length;

  const sessionRes = await betaClient.post<{ uploadUrl: string }>(
    `/me/todo/lists/${listId}/tasks/${taskId}/attachments/createUploadSession`,
    {
      attachmentInfo: {
        attachmentType: 'file',
        name: attachment.name,
        size: fileSize,
      },
    }
  );

  const uploadUrl = sessionRes.data.uploadUrl;

  const CHUNK_SIZE = 4 * 1024 * 1024;
  let offset = 0;

  while (offset < fileSize) {
    const chunkEnd = Math.min(offset + CHUNK_SIZE, fileSize);
    const chunk = contentBytes.slice(offset, chunkEnd);

    const uploadRes = await fetch(`${uploadUrl}/content`, {
      method: 'PUT',
      headers: {
        Authorization: ['Bearer', await getAccessToken()].join(' '),
        'Content-Type': 'application/octet-stream',
        'Content-Range': `bytes ${offset}-${chunkEnd - 1}/${fileSize}`,
      },
      body: chunk,
    });
    if (!uploadRes.ok) {
      const errorBody = await uploadRes.text();
      let message: string | undefined;
      try {
        message = JSON.parse(errorBody)?.error?.message;
      } catch {
        message = undefined;
      }
      throw new AppError(ErrorCodes.GRAPH_ERROR, message || uploadRes.statusText || `HTTP ${uploadRes.status}`);
    }

    const location = uploadRes.headers.get('location');
    if (uploadRes.status === 201 && location) {
      const attachmentId = location.split('/').pop();
      return {
        id: attachmentId || '',
        name: attachment.name,
        size: fileSize,
        contentType: attachment.contentType,
      };
    }

    offset = chunkEnd;
  }

  throw new AppError(ErrorCodes.GRAPH_ERROR, 'Failed to upload attachment');
}

export async function getTasksAcrossLists(): Promise<TodoTask[]> {
  const client = createClient();
  const lists = await fetchPaged<TodoList, TodoList>(client, '/me/todo/lists', (list) => list);
  if (lists.length === 0) return [];

  const results = await mapConcurrent(lists, TASK_FETCH_BATCH_SIZE, (list) => {
    return fetchPaged<GraphTask, TodoTask>(client, `/me/todo/lists/${list.id}/tasks`, (task) =>
      mapTask(task, list.displayName, list.id),
    );
  });

  return results.flat();
}
