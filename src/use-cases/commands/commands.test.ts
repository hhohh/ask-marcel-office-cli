import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../../domain/access-token.ts';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { AuthManager } from '../../infra/auth.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { createGraphClient } from '../../infra/graph-client.ts';
import * as downloadDriveItemAsMarkdown from './download-drive-item-as-markdown.ts';
import * as downloadDriveItemAsPdf from './download-drive-item-as-pdf.ts';
import * as downloadDriveItemVersionAsMarkdown from './download-drive-item-version-as-markdown.ts';
import * as downloadDriveItemVersionAsPdf from './download-drive-item-version-as-pdf.ts';
import * as downloadDriveItemVersionContent from './download-drive-item-version-content.ts';
import * as downloadOnedriveFileContent from './download-onedrive-file-content.ts';
import * as getCalendarEvent from './get-calendar-event.ts';
import * as getCalendarView from './get-calendar-view.ts';
import * as getCurrentUser from './get-current-user.ts';
import * as getDriveDelta from './get-drive-delta.ts';
import * as getDriveItem from './get-drive-item.ts';
import * as getDriveRootItem from './get-drive-root-item.ts';
import * as getExcelRange from './get-excel-range.ts';
import * as getExcelTable from './get-excel-table.ts';
import * as getMailAttachment from './get-mail-attachment.ts';
import * as getMailMessage from './get-mail-message.ts';
import * as getMailboxSettings from './get-mailbox-settings.ts';
import * as getMyProfilePhoto from './get-my-profile-photo.ts';
import * as getOnenotePageAsMarkdown from './get-onenote-page-as-markdown.ts';
import * as getOnenotePageContent from './get-onenote-page-content.ts';
import * as getPlannerBucket from './get-planner-bucket.ts';
import * as getPlannerPlan from './get-planner-plan.ts';
import * as getPlannerTaskDetails from './get-planner-task-details.ts';
import * as getPlannerTask from './get-planner-task.ts';
import * as getSharepointSiteByPath from './get-sharepoint-site-by-path.ts';
import * as getSharepointSiteDriveById from './get-sharepoint-site-drive-by-id.ts';
import * as getSharepointSiteListItem from './get-sharepoint-site-list-item.ts';
import * as getSharepointSiteList from './get-sharepoint-site-list.ts';
import * as getSharepointSite from './get-sharepoint-site.ts';
import * as getSpecificCalendarEvent from './get-specific-calendar-event.ts';
import * as getSpecificCalendarView from './get-specific-calendar-view.ts';
import * as getTeamChannel from './get-team-channel.ts';
import * as getTeam from './get-team.ts';
import * as getTodoTask from './get-todo-task.ts';
import * as listAllOnenoteSections from './list-all-onenote-sections.ts';
import * as listCalendarEventInstances from './list-calendar-event-instances.ts';
import * as listCalendarEventsDelta from './list-calendar-events-delta.ts';
import * as listCalendarEvents from './list-calendar-events.ts';
import * as listCalendarViewDelta from './list-calendar-view-delta.ts';
import * as listCalendars from './list-calendars.ts';
import * as listChatMembers from './list-chat-members.ts';
import * as listDriveItemPermissions from './list-drive-item-permissions.ts';
import * as listDriveItemVersions from './list-drive-item-versions.ts';
import * as listDrives from './list-drives.ts';
import * as listExcelTableRows from './list-excel-table-rows.ts';
import * as listExcelTables from './list-excel-tables.ts';
import * as listExcelWorksheets from './list-excel-worksheets.ts';
import * as listFolderFiles from './list-folder-files.ts';
import * as listIncompletePlannerTasks from './list-incomplete-planner-tasks.ts';
import * as listIncompleteTodoTasks from './list-incomplete-todo-tasks.ts';
import * as listJoinedTeams from './list-joined-teams.ts';
import * as listMailAttachments from './list-mail-attachments.ts';
import * as listMailChildFolders from './list-mail-child-folders.ts';
import * as listMailFolderMessages from './list-mail-folder-messages.ts';
import * as listMailFolders from './list-mail-folders.ts';
import * as listMailMessages from './list-mail-messages.ts';
import * as listMailRules from './list-mail-rules.ts';
import * as listOnenoteNotebookSections from './list-onenote-notebook-sections.ts';
import * as listOnenoteNotebooks from './list-onenote-notebooks.ts';
import * as listOnenoteSectionPages from './list-onenote-section-pages.ts';
import * as listPlanBuckets from './list-plan-buckets.ts';
import * as listPlanTasks from './list-plan-tasks.ts';
import * as listPlannerPlans from './list-planner-plans.ts';
import * as listPlannerTasks from './list-planner-tasks.ts';
import * as listSharepointSiteDrives from './list-sharepoint-site-drives.ts';
import * as listSharepointSiteListItems from './list-sharepoint-site-list-items.ts';
import * as listSharepointSiteLists from './list-sharepoint-site-lists.ts';
import * as listSpecificCalendarEvents from './list-specific-calendar-events.ts';
import * as listTeamChannels from './list-team-channels.ts';
import * as listTodoLinkedResources from './list-todo-linked-resources.ts';
import * as listTodoTaskLists from './list-todo-task-lists.ts';
import * as listTodoTasks from './list-todo-tasks.ts';
import * as nextPage from './next-page.ts';
import * as searchMailMessages from './search-mail-messages.ts';
import * as searchMyDocuments from './search-my-documents.ts';
import * as searchOnedriveFiles from './search-onedrive-files.ts';
import * as searchOnenotePages from './search-onenote-pages.ts';
import * as searchSharepointSitesByName from './search-sharepoint-sites-by-name.ts';
import * as searchSharepointSites from './search-sharepoint-sites.ts';

const cmdMap: Record<string, { execute: typeof listDrives.execute }> = {
  'list-drives': listDrives,
  'get-drive-root-item': getDriveRootItem,
  'list-folder-files': listFolderFiles,
  'download-onedrive-file-content': downloadOnedriveFileContent,
  'get-drive-item': getDriveItem,
  'list-drive-item-permissions': listDriveItemPermissions,
  'list-drive-item-versions': listDriveItemVersions,
  'download-drive-item-version-content': downloadDriveItemVersionContent,
  'download-drive-item-as-pdf': downloadDriveItemAsPdf,
  'download-drive-item-as-markdown': downloadDriveItemAsMarkdown,
  'download-drive-item-version-as-pdf': downloadDriveItemVersionAsPdf,
  'download-drive-item-version-as-markdown': downloadDriveItemVersionAsMarkdown,
  'search-onedrive-files': searchOnedriveFiles,
  'search-my-documents': searchMyDocuments,
  'get-excel-range': getExcelRange,
  'list-excel-worksheets': listExcelWorksheets,
  'list-excel-tables': listExcelTables,
  'get-excel-table': getExcelTable,
  'list-excel-table-rows': listExcelTableRows,
  'get-drive-delta': getDriveDelta,
  'search-sharepoint-sites': searchSharepointSites,
  'search-sharepoint-sites-by-name': searchSharepointSitesByName,
  'get-sharepoint-site': getSharepointSite,
  'list-sharepoint-site-drives': listSharepointSiteDrives,
  'get-sharepoint-site-drive-by-id': getSharepointSiteDriveById,
  'list-sharepoint-site-lists': listSharepointSiteLists,
  'get-sharepoint-site-list': getSharepointSiteList,
  'list-sharepoint-site-list-items': listSharepointSiteListItems,
  'get-sharepoint-site-list-item': getSharepointSiteListItem,
  'get-sharepoint-site-by-path': getSharepointSiteByPath,
  'list-todo-task-lists': listTodoTaskLists,
  'list-todo-tasks': listTodoTasks,
  'list-incomplete-todo-tasks': listIncompleteTodoTasks,
  'get-todo-task': getTodoTask,
  'list-todo-linked-resources': listTodoLinkedResources,
  'list-planner-plans': listPlannerPlans,
  'list-planner-tasks': listPlannerTasks,
  'list-incomplete-planner-tasks': listIncompletePlannerTasks,
  'get-planner-plan': getPlannerPlan,
  'list-plan-tasks': listPlanTasks,
  'get-planner-task': getPlannerTask,
  'get-planner-task-details': getPlannerTaskDetails,
  'list-plan-buckets': listPlanBuckets,
  'get-planner-bucket': getPlannerBucket,
  'list-mail-messages': listMailMessages,
  'list-mail-folders': listMailFolders,
  'list-mail-child-folders': listMailChildFolders,
  'list-mail-folder-messages': listMailFolderMessages,
  'get-mail-message': getMailMessage,
  'list-mail-attachments': listMailAttachments,
  'get-mail-attachment': getMailAttachment,
  'list-mail-rules': listMailRules,
  'get-mailbox-settings': getMailboxSettings,
  'search-mail-messages': searchMailMessages,
  'list-onenote-notebooks': listOnenoteNotebooks,
  'list-onenote-notebook-sections': listOnenoteNotebookSections,
  'list-all-onenote-sections': listAllOnenoteSections,
  'list-onenote-section-pages': listOnenoteSectionPages,
  'get-onenote-page-content': getOnenotePageContent,
  'get-onenote-page-as-markdown': getOnenotePageAsMarkdown,
  'search-onenote-pages': searchOnenotePages,
  'get-current-user': getCurrentUser,
  'get-my-profile-photo': getMyProfilePhoto,
  'list-calendar-events': listCalendarEvents,
  'get-calendar-event': getCalendarEvent,
  'list-specific-calendar-events': listSpecificCalendarEvents,
  'get-specific-calendar-event': getSpecificCalendarEvent,
  'get-calendar-view': getCalendarView,
  'get-specific-calendar-view': getSpecificCalendarView,
  'list-calendar-event-instances': listCalendarEventInstances,
  'list-calendars': listCalendars,
  'list-calendar-events-delta': listCalendarEventsDelta,
  'list-calendar-view-delta': listCalendarViewDelta,
  'list-chat-members': listChatMembers,
  'list-joined-teams': listJoinedTeams,
  'get-team': getTeam,
  'list-team-channels': listTeamChannels,
  'get-team-channel': getTeamChannel,
  'next-page': nextPage,
};

const fakeAuth = (): AuthManager => ({ getAccessToken: async () => ok(accessTokenUnsafe('test-token')), logout: async () => ok(undefined) });

type FakeFetch = ((url: string, init?: RequestInit) => Promise<Response>) & { lastUrl: string | null; lastBody: string | null };

const fakeFetch = (body: unknown): FakeFetch => {
  let lastUrl: string | null = null;
  let lastBody: string | null = null;
  const fn = async (url: string, init?: RequestInit): Promise<Response> => {
    lastUrl = url;
    lastBody = typeof init?.body === 'string' ? init.body : null;
    return new Response(JSON.stringify(body), { headers: { 'content-type': 'application/json' } });
  };
  Object.defineProperty(fn, 'lastUrl', { get: () => lastUrl });
  Object.defineProperty(fn, 'lastBody', { get: () => lastBody });
  return fn as FakeFetch;
};

const callCommand = async (name: string, params: Record<string, string>, responseBody: unknown): Promise<Result<unknown, GraphError>> => {
  const cmd = cmdMap[name];
  if (!cmd) throw new Error(`command not found: ${name}`);
  const fetchFn = fakeFetch(responseBody);
  const graph = createGraphClient(fakeAuth(), fetchFn);
  return cmd.execute(graph, params);
};

const capturedUrl = async (name: string, params: Record<string, string>): Promise<string> => {
  const cmd = cmdMap[name];
  if (!cmd) throw new Error(`command not found: ${name}`);
  const fetchFn = fakeFetch({ ok: true });
  const graph = createGraphClient(fakeAuth(), fetchFn);
  await cmd.execute(graph, params);
  return fetchFn.lastUrl ?? '';
};

/**
 * Stateful fetch fake for multi-step commands (per the FIX #4 design):
 * each handler matches the FIRST entry whose `urlPrefix` is a prefix
 * of the request URL AND whose optional `method` matches. The matched
 * handler is then CONSUMED — subsequent calls fall through to the
 * next available handler. Unmatched calls throw a clear error rather
 * than silently mismatching. `response` may be a static `Response` or
 * a thunk so each call gets a fresh body / headers.
 */
type StagedHandler = {
  readonly urlPrefix: string;
  readonly method?: string;
  readonly response: Response | (() => Response);
};

const stagedFetch = (handlers: ReadonlyArray<StagedHandler>): ((url: string, init?: RequestInit) => Promise<Response>) => {
  const queue = handlers.map((h) => ({ ...h, consumed: false }));
  return async (url, init) => {
    const requestUrl = url.split('?')[0] ?? url;
    const requestUrlWithQuery = url;
    const method = init?.method ?? 'GET';
    const idx = queue.findIndex((h) => {
      if (h.consumed) return false;
      if (h.method !== undefined && h.method !== method) return false;
      const prefix = h.urlPrefix.split('?')[0] ?? h.urlPrefix;
      const queryNeeded = h.urlPrefix.includes('?');
      if (queryNeeded) return requestUrlWithQuery.startsWith(h.urlPrefix);
      return requestUrl.startsWith(prefix);
    });
    if (idx === -1) throw new Error(`stagedFetch: unexpected ${method} ${url}`);
    const handler = queue[idx];
    if (!handler) throw new Error(`stagedFetch: corrupt queue at ${idx}`);
    handler.consumed = true;
    return typeof handler.response === 'function' ? handler.response() : handler.response.clone();
  };
};

describe('commands', () => {
  it('list-drives returns drives', async () => {
    const result = await callCommand('list-drives', {}, { value: [{ id: 'd1' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ id: 'd1' }] });
  });

  it('get-drive-root-item returns root item', async () => {
    const result = await callCommand('get-drive-root-item', { driveId: 'drive-1' }, { id: 'root', name: 'Root' });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ id: 'root', name: 'Root' });
  });

  it('list-folder-files returns children', async () => {
    const result = await callCommand('list-folder-files', { driveId: 'd1', itemId: 'i1' }, { value: [{ name: 'file.txt' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ name: 'file.txt' }] });
  });

  it('get-drive-item returns metadata', async () => {
    const result = await callCommand('get-drive-item', { driveId: 'd1', itemId: 'i1' }, { name: 'doc.xlsx', size: 1024 });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ name: 'doc.xlsx', size: 1024 });
  });

  it('search-onedrive-files searches with query', async () => {
    const result = await callCommand('search-onedrive-files', { driveId: 'd1', query: 'report' }, { value: [{ name: 'report.xlsx' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ name: 'report.xlsx' }] });
  });

  it('search-my-documents searches the user’s default OneDrive', async () => {
    const result = await callCommand('search-my-documents', { query: 'budget' }, { value: [{ name: 'budget.xlsx' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ name: 'budget.xlsx' }] });
  });

  it('search-mail-messages searches the mailbox with $search', async () => {
    const result = await callCommand('search-mail-messages', { query: 'invoice' }, { value: [{ subject: 'invoice 042' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ subject: 'invoice 042' }] });
  });

  it('search-onenote-pages filters OneNote pages by title (Graph removed full-text ?search= from v1.0)', async () => {
    const result = await callCommand('search-onenote-pages', { titleSubstring: 'meeting notes' }, { value: [{ title: 'Meeting notes 2026-04-30' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ title: 'Meeting notes 2026-04-30' }] });
  });

  it('search-sharepoint-sites-by-name searches sites with the search query parameter', async () => {
    const result = await callCommand('search-sharepoint-sites-by-name', { query: 'marketing' }, { value: [{ displayName: 'Marketing site' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ displayName: 'Marketing site' }] });
  });

  it('list-incomplete-planner-tasks filters Planner tasks by percentComplete ne 100', async () => {
    const result = await callCommand('list-incomplete-planner-tasks', {}, { value: [{ id: 'pt1', percentComplete: 0 }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ id: 'pt1', percentComplete: 0 }] });
  });

  it('list-incomplete-todo-tasks filters To Do tasks by status ne completed within a list', async () => {
    const result = await callCommand('list-incomplete-todo-tasks', { todoTaskListId: 'tl1' }, { value: [{ id: 'tt1', status: 'inProgress' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ id: 'tt1', status: 'inProgress' }] });
  });

  it('next-page strips the Graph v1.0 prefix and GETs the rest of the supplied URL', async () => {
    const result = await callCommand('next-page', { url: 'https://graph.microsoft.com/v1.0/me/messages?$skip=10' }, { value: [{ subject: 'page 2' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ subject: 'page 2' }] });
  });

  it('next-page rejects URLs that do not start with the Graph v1.0 prefix', async () => {
    const cmd = cmdMap['next-page'];
    if (!cmd) throw new Error('next-page not registered');
    const fetchFn = fakeFetch({ ok: true });
    const graph = createGraphClient(fakeAuth(), fetchFn);
    try {
      await cmd.execute(graph, { url: 'https://example.com/something' });
      throw new Error('should have rejected');
    } catch (e) {
      expect((e as Error).message).toContain('validation failed:');
    }
  });

  it('get-excel-range returns cell values', async () => {
    const result = await callCommand(
      'get-excel-range',
      { driveId: 'd1', itemId: 'i1', worksheetId: 'ws1', address: 'A1:B2' },
      {
        values: [
          ['a', 'b'],
          ['c', 'd'],
        ],
      }
    );
    expect(result.ok).toBe(true);
    if (result.ok)
      expect(result.value).toEqual({
        values: [
          ['a', 'b'],
          ['c', 'd'],
        ],
      });
  });

  it('list-excel-worksheets returns worksheets', async () => {
    const result = await callCommand('list-excel-worksheets', { driveId: 'd1', itemId: 'i1' }, { value: [{ name: 'Sheet1' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ name: 'Sheet1' }] });
  });

  it('list-excel-tables returns tables', async () => {
    const result = await callCommand('list-excel-tables', { driveId: 'd1', itemId: 'i1' }, { value: [{ name: 'Table1' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ name: 'Table1' }] });
  });

  it('get-excel-table returns table details', async () => {
    const result = await callCommand('get-excel-table', { driveId: 'd1', itemId: 'i1', tableId: 't1' }, { name: 'Table1', showHeaders: true });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ name: 'Table1', showHeaders: true });
  });

  it('list-excel-table-rows returns rows', async () => {
    const result = await callCommand('list-excel-table-rows', { driveId: 'd1', itemId: 'i1', tableId: 't1' }, { value: [{ index: 0, values: ['a', 'b'] }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ index: 0, values: ['a', 'b'] }] });
  });

  it('get-drive-delta returns changes', async () => {
    const result = await callCommand('get-drive-delta', { driveId: 'd1', itemId: 'i1' }, { value: [{ id: 'new-file' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ id: 'new-file' }] });
  });

  it('list-drive-item-permissions returns permissions', async () => {
    const result = await callCommand('list-drive-item-permissions', { driveId: 'd1', itemId: 'i1' }, { value: [{ roles: ['read'] }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ roles: ['read'] }] });
  });

  it('list-drive-item-versions returns versions', async () => {
    const result = await callCommand('list-drive-item-versions', { driveId: 'd1', itemId: 'i1' }, { value: [{ id: 'v1' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ id: 'v1' }] });
  });

  it('download-onedrive-file-content returns content info', async () => {
    const result = await callCommand('download-onedrive-file-content', { driveId: 'd1', itemId: 'i1' }, { '@microsoft.graph.downloadUrl': 'https://...' });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ '@microsoft.graph.downloadUrl': 'https://...' });
  });

  it('download-drive-item-version-content returns content for a historical version', async () => {
    const result = await callCommand(
      'download-drive-item-version-content',
      { driveId: 'd1', itemId: 'i1', versionId: '3.0' },
      { '@microsoft.graph.downloadUrl': 'https://cdn.example/v3' }
    );
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ '@microsoft.graph.downloadUrl': 'https://cdn.example/v3' });
  });

  it('download-drive-item-as-pdf converts an Office source via Graph ?format=pdf after the metadata pre-check', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1', method: 'GET', response: Response.json({ name: 'q3.docx', size: 9 }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/content?format=pdf',
        method: 'GET',
        response: () => Response.json({ '@microsoft.graph.downloadUrl': 'https://cdn.example/q3.pdf' }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1' });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ '@microsoft.graph.downloadUrl': 'https://cdn.example/q3.pdf' });
  });

  it('download-drive-item-as-pdf short-circuits to a raw bytes download for plain-text source extensions', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText', method: 'GET', response: Response.json({ name: 'README.md', size: 4 }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText/content',
        method: 'GET',
        response: () => new Response('hi', { status: 200, headers: { 'content-type': 'text/markdown' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iText' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toBe('hi');
    }
  });

  it('download-drive-item-as-markdown converts Office HTML to markdown via the pipeline', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1', method: 'GET', response: Response.json({ name: 'q3.docx' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/content?format=html',
        method: 'GET',
        response: () => new Response('<h1>Q3</h1>', { status: 200, headers: { 'content-type': 'text/html' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toContain('# Q3');
    }
  });

  it('download-drive-item-as-markdown short-circuits to raw download for plain-text source extensions', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText', method: 'GET', response: Response.json({ name: 'notes.txt' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText/content',
        method: 'GET',
        response: () => new Response('plain', { status: 200, headers: { 'content-type': 'text/plain' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iText' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toBe('plain');
    }
  });

  it('download-drive-item-as-pdf propagates an err from the metadata pre-fetch unchanged', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iMissing',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { message: 'not found' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iMissing' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
    }
  });

  it('download-drive-item-version-as-pdf converts a non-current version through Graph ?format=pdf', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1', method: 'GET', response: Response.json({ name: 'budget.xlsx' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/versions/3.0/content?format=pdf',
        method: 'GET',
        response: () => Response.json({ '@microsoft.graph.downloadUrl': 'https://cdn.example/v3.pdf' }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-version-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', versionId: '3.0' });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ '@microsoft.graph.downloadUrl': 'https://cdn.example/v3.pdf' });
  });

  it('download-drive-item-version-as-markdown converts a non-current version HTML to markdown', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1', method: 'GET', response: Response.json({ name: 'budget.xlsx' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/versions/3.0/content?format=html',
        method: 'GET',
        response: () => new Response('<h2>v3</h2>', { status: 200, headers: { 'content-type': 'text/html' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-version-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', versionId: '3.0' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('## v3');
    }
  });

  it('download-drive-item-version-as-markdown short-circuits to raw download for plain-text source extensions', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText', method: 'GET', response: Response.json({ name: 'notes.md' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText/versions/2.0/content',
        method: 'GET',
        response: () => new Response('# v2', { status: 200, headers: { 'content-type': 'text/markdown' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-version-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iText', versionId: '2.0' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toBe('# v2');
    }
  });

  it('download-drive-item-version-as-pdf short-circuits to raw download for plain-text source extensions', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText', method: 'GET', response: Response.json({ name: 'log.log' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText/versions/2.0/content',
        method: 'GET',
        response: () => new Response('line', { status: 200, headers: { 'content-type': 'text/plain' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-version-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iText', versionId: '2.0' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toBe('line');
    }
  });

  it('download-drive-item-version-as-pdf propagates an err from the metadata pre-fetch unchanged', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iMissing',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-version-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iMissing', versionId: '2.0' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
    }
  });

  it('download-drive-item-version-as-markdown propagates an err from the metadata pre-fetch unchanged', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iMissing',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-version-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iMissing', versionId: '2.0' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
    }
  });

  it('download-drive-item-as-markdown propagates an err from the metadata pre-fetch unchanged', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iMissing',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iMissing' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
    }
  });

  it('get-onenote-page-as-markdown turns Graph-returned HTML into a markdown envelope', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/onenote/pages/p1/content',
        method: 'GET',
        response: () => new Response('<h1>Meeting</h1>', { status: 200, headers: { 'content-type': 'text/html' } }),
      },
    ]);
    const cmd = cmdMap['get-onenote-page-as-markdown'];
    if (!cmd) throw new Error('get-onenote-page-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { onenotePageId: 'p1' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toContain('# Meeting');
    }
  });
});

type CommandFixture = { readonly name: string; readonly params: Record<string, string> };

const allCommandFixtures: CommandFixture[] = [
  { name: 'list-drives', params: {} },
  { name: 'get-drive-root-item', params: { driveId: 'd1' } },
  { name: 'list-folder-files', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'download-onedrive-file-content', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'get-drive-item', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'list-drive-item-permissions', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'list-drive-item-versions', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'download-drive-item-version-content', params: { driveId: 'd1', itemId: 'i1', versionId: '3.0' } },
  { name: 'download-drive-item-as-pdf', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'download-drive-item-version-as-pdf', params: { driveId: 'd1', itemId: 'i1', versionId: '3.0' } },
  { name: 'search-onedrive-files', params: { driveId: 'd1', query: 'report' } },
  { name: 'search-my-documents', params: { query: 'budget' } },
  { name: 'get-excel-range', params: { driveId: 'd1', itemId: 'i1', worksheetId: 'ws1', address: 'A1' } },
  { name: 'list-excel-worksheets', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'list-excel-tables', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'get-excel-table', params: { driveId: 'd1', itemId: 'i1', tableId: 't1' } },
  { name: 'list-excel-table-rows', params: { driveId: 'd1', itemId: 'i1', tableId: 't1' } },
  { name: 'get-drive-delta', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'search-sharepoint-sites', params: {} },
  { name: 'search-sharepoint-sites-by-name', params: { query: 'marketing' } },
  { name: 'get-sharepoint-site', params: { siteId: 's1' } },
  { name: 'list-sharepoint-site-drives', params: { siteId: 's1' } },
  { name: 'get-sharepoint-site-drive-by-id', params: { siteId: 's1', driveId: 'd1' } },
  { name: 'list-sharepoint-site-lists', params: { siteId: 's1' } },
  { name: 'get-sharepoint-site-list', params: { siteId: 's1', listId: 'l1' } },
  { name: 'list-sharepoint-site-list-items', params: { siteId: 's1', listId: 'l1' } },
  { name: 'get-sharepoint-site-list-item', params: { siteId: 's1', listId: 'l1', listItemId: 'li1' } },
  { name: 'get-sharepoint-site-by-path', params: { hostname: 'contoso.sharepoint.com', path: '/sites/Marketing' } },
  { name: 'list-todo-task-lists', params: {} },
  { name: 'list-todo-tasks', params: { todoTaskListId: 'tl1' } },
  { name: 'list-incomplete-todo-tasks', params: { todoTaskListId: 'tl1' } },
  { name: 'get-todo-task', params: { todoTaskListId: 'tl1', todoTaskId: 't1' } },
  { name: 'list-todo-linked-resources', params: { todoTaskListId: 'tl1', todoTaskId: 't1' } },
  { name: 'list-planner-plans', params: {} },
  { name: 'list-planner-tasks', params: {} },
  { name: 'list-incomplete-planner-tasks', params: {} },
  { name: 'get-planner-plan', params: { plannerPlanId: 'p1' } },
  { name: 'list-plan-tasks', params: { plannerPlanId: 'p1' } },
  { name: 'get-planner-task', params: { plannerTaskId: 't1' } },
  { name: 'get-planner-task-details', params: { plannerTaskId: 't1' } },
  { name: 'list-plan-buckets', params: { plannerPlanId: 'p1' } },
  { name: 'get-planner-bucket', params: { plannerBucketId: 'b1' } },
  { name: 'list-mail-messages', params: {} },
  { name: 'list-mail-folders', params: {} },
  { name: 'list-mail-child-folders', params: { mailFolderId: 'f1' } },
  { name: 'list-mail-folder-messages', params: { mailFolderId: 'f1' } },
  { name: 'get-mail-message', params: { messageId: 'm1' } },
  { name: 'list-mail-attachments', params: { messageId: 'm1' } },
  { name: 'get-mail-attachment', params: { messageId: 'm1', attachmentId: 'a1' } },
  { name: 'list-mail-rules', params: { mailFolderId: 'f1' } },
  { name: 'get-mailbox-settings', params: {} },
  { name: 'search-mail-messages', params: { query: 'invoice' } },
  { name: 'list-onenote-notebooks', params: {} },
  { name: 'list-onenote-notebook-sections', params: { notebookId: 'n1' } },
  { name: 'list-all-onenote-sections', params: {} },
  { name: 'list-onenote-section-pages', params: { onenoteSectionId: 's1' } },
  { name: 'get-onenote-page-content', params: { onenotePageId: 'p1' } },
  { name: 'search-onenote-pages', params: { titleSubstring: 'meeting' } },
  { name: 'get-current-user', params: {} },
  { name: 'get-my-profile-photo', params: {} },
  { name: 'list-calendar-events', params: {} },
  { name: 'get-calendar-event', params: { eventId: 'e1' } },
  { name: 'list-specific-calendar-events', params: { calendarId: 'c1' } },
  { name: 'get-specific-calendar-event', params: { calendarId: 'c1', eventId: 'e1' } },
  { name: 'get-calendar-view', params: { startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' } },
  { name: 'get-specific-calendar-view', params: { calendarId: 'c1', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' } },
  { name: 'list-calendar-event-instances', params: { calendarId: 'c1', eventId: 'e1', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' } },
  { name: 'list-calendars', params: {} },
  { name: 'list-calendar-events-delta', params: {} },
  { name: 'list-calendar-view-delta', params: { startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' } },
  { name: 'list-chat-members', params: { chatId: 'ch1' } },
  { name: 'list-joined-teams', params: {} },
  { name: 'get-team', params: { teamId: 'tm1' } },
  { name: 'list-team-channels', params: { teamId: 'tm1' } },
  { name: 'get-team-channel', params: { teamId: 'tm1', channelId: 'ch1' } },
  { name: 'next-page', params: { url: 'https://graph.microsoft.com/v1.0/me/messages?$skip=10' } },
];

describe('all commands schema acceptance', () => {
  it.each(allCommandFixtures)('$name accepts valid params', async ({ name, params }) => {
    const result = await callCommand(name, params, { ok: true });
    expect(result.ok).toBe(true);
  });
});

describe('command schema rejection', () => {
  const rejectCases: Array<{ name: string; params: Record<string, string> }> = [
    { name: 'get-drive-root-item', params: {} },
    { name: 'search-onedrive-files', params: { driveId: 'd1' } },
    { name: 'get-sharepoint-site', params: {} },
    { name: 'get-sharepoint-site-by-path', params: {} },
    { name: 'get-sharepoint-site-by-path', params: { hostname: 'contoso.sharepoint.com', path: 'sites/Marketing' } },
    { name: 'get-todo-task', params: { todoTaskListId: 'tl1' } },
    { name: 'get-planner-plan', params: {} },
    { name: 'get-mail-message', params: {} },
    { name: 'list-mail-child-folders', params: {} },
    { name: 'get-onenote-page-content', params: {} },
    { name: 'get-calendar-event', params: {} },
    { name: 'get-specific-calendar-event', params: { calendarId: 'c1' } },
    { name: 'list-team-channels', params: {} },
    { name: 'get-team-channel', params: { teamId: 'tm1' } },
    { name: 'download-onedrive-file-content', params: { driveId: 'd1' } },
    { name: 'download-drive-item-version-content', params: { driveId: 'd1', itemId: 'i1' } },
    { name: 'download-drive-item-as-pdf', params: { driveId: 'd1' } },
    { name: 'download-drive-item-as-markdown', params: { driveId: 'd1' } },
    { name: 'download-drive-item-version-as-pdf', params: { driveId: 'd1', itemId: 'i1' } },
    { name: 'download-drive-item-version-as-markdown', params: { driveId: 'd1', itemId: 'i1' } },
    { name: 'get-onenote-page-as-markdown', params: {} },
    { name: 'search-mail-messages', params: {} },
    { name: 'search-my-documents', params: {} },
    { name: 'get-calendar-view', params: {} },
    { name: 'list-calendar-view-delta', params: {} },
    { name: 'get-specific-calendar-view', params: { calendarId: 'c1' } },
    { name: 'list-calendar-event-instances', params: { calendarId: 'c1', eventId: 'e1' } },
    { name: 'search-onenote-pages', params: {} },
    { name: 'search-sharepoint-sites-by-name', params: {} },
    { name: 'list-incomplete-todo-tasks', params: {} },
    { name: 'next-page', params: {} },
    { name: 'next-page', params: { url: 'https://example.com/foo' } },
  ];

  it.each(rejectCases)('$name rejects missing required params', async ({ name, params }) => {
    const cmd = cmdMap[name];
    const fetchFn = fakeFetch({ ok: true });
    const graph = createGraphClient(fakeAuth(), fetchFn);
    try {
      await cmd.execute(graph, params);
    } catch (e) {
      expect((e as Error).message).toContain('validation failed:');
    }
  });
});

const pathFixtures: Array<{ name: string; params: Record<string, string>; expectedPath: string }> = [
  { name: 'list-drives', params: {}, expectedPath: '/me/drives' },
  { name: 'get-drive-root-item', params: { driveId: 'd1' }, expectedPath: '/drives/d1/root' },
  { name: 'list-folder-files', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/children' },
  { name: 'download-onedrive-file-content', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/content' },
  { name: 'get-drive-item', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1' },
  { name: 'list-drive-item-permissions', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/permissions' },
  { name: 'list-drive-item-versions', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/versions' },
  {
    name: 'download-drive-item-version-content',
    params: { driveId: 'd1', itemId: 'i1', versionId: '3.0' },
    expectedPath: '/drives/d1/items/i1/versions/3.0/content',
  },
  {
    name: 'download-drive-item-as-pdf',
    params: { driveId: 'd1', itemId: 'i1' },
    expectedPath: '/drives/d1/items/i1/content?format=pdf',
  },
  {
    name: 'download-drive-item-as-markdown',
    params: { driveId: 'd1', itemId: 'i1' },
    expectedPath: '/drives/d1/items/i1/content?format=html',
  },
  {
    name: 'download-drive-item-version-as-pdf',
    params: { driveId: 'd1', itemId: 'i1', versionId: '3.0' },
    expectedPath: '/drives/d1/items/i1/versions/3.0/content?format=pdf',
  },
  {
    name: 'download-drive-item-version-as-markdown',
    params: { driveId: 'd1', itemId: 'i1', versionId: '3.0' },
    expectedPath: '/drives/d1/items/i1/versions/3.0/content?format=html',
  },
  { name: 'search-onedrive-files', params: { driveId: 'd1', query: 'report' }, expectedPath: "/drives/d1/search(q='report')" },
  { name: 'search-my-documents', params: { query: 'budget' }, expectedPath: "/me/drive/search(q='budget')" },
  {
    name: 'get-excel-range',
    params: { driveId: 'd1', itemId: 'i1', worksheetId: 'ws1', address: 'A1' },
    expectedPath: "/drives/d1/items/i1/workbook/worksheets/ws1/range(address='A1')",
  },
  { name: 'list-excel-worksheets', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/workbook/worksheets' },
  { name: 'list-excel-tables', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/workbook/tables' },
  { name: 'get-excel-table', params: { driveId: 'd1', itemId: 'i1', tableId: 't1' }, expectedPath: '/drives/d1/items/i1/workbook/tables/t1' },
  { name: 'list-excel-table-rows', params: { driveId: 'd1', itemId: 'i1', tableId: 't1' }, expectedPath: '/drives/d1/items/i1/workbook/tables/t1/rows' },
  { name: 'get-drive-delta', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/delta()' },
  { name: 'search-sharepoint-sites', params: {}, expectedPath: '/sites' },
  { name: 'search-sharepoint-sites-by-name', params: { query: 'marketing' }, expectedPath: '/sites?search=marketing' },
  { name: 'get-sharepoint-site', params: { siteId: 's1' }, expectedPath: '/sites/s1' },
  { name: 'list-sharepoint-site-drives', params: { siteId: 's1' }, expectedPath: '/sites/s1/drives' },
  { name: 'get-sharepoint-site-drive-by-id', params: { siteId: 's1', driveId: 'd1' }, expectedPath: '/sites/s1/drives/d1' },
  { name: 'list-sharepoint-site-lists', params: { siteId: 's1' }, expectedPath: '/sites/s1/lists' },
  { name: 'get-sharepoint-site-list', params: { siteId: 's1', listId: 'l1' }, expectedPath: '/sites/s1/lists/l1' },
  { name: 'list-sharepoint-site-list-items', params: { siteId: 's1', listId: 'l1' }, expectedPath: '/sites/s1/lists/l1/items' },
  { name: 'get-sharepoint-site-list-item', params: { siteId: 's1', listId: 'l1', listItemId: 'li1' }, expectedPath: '/sites/s1/lists/l1/items/li1' },
  {
    name: 'get-sharepoint-site-by-path',
    params: { hostname: 'contoso.sharepoint.com', path: '/sites/Marketing' },
    expectedPath: '/sites/contoso.sharepoint.com:/sites/Marketing',
  },
  { name: 'list-todo-task-lists', params: {}, expectedPath: '/me/todo/lists' },
  { name: 'list-todo-tasks', params: { todoTaskListId: 'tl1' }, expectedPath: '/me/todo/lists/tl1/tasks' },
  { name: 'list-incomplete-todo-tasks', params: { todoTaskListId: 'tl1' }, expectedPath: "/me/todo/lists/tl1/tasks?$filter=status ne 'completed'" },
  { name: 'get-todo-task', params: { todoTaskListId: 'tl1', todoTaskId: 't1' }, expectedPath: '/me/todo/lists/tl1/tasks/t1' },
  { name: 'list-todo-linked-resources', params: { todoTaskListId: 'tl1', todoTaskId: 't1' }, expectedPath: '/me/todo/lists/tl1/tasks/t1/linkedResources' },
  { name: 'list-planner-plans', params: {}, expectedPath: '/me/planner/plans' },
  { name: 'list-planner-tasks', params: {}, expectedPath: '/me/planner/tasks' },
  { name: 'list-incomplete-planner-tasks', params: {}, expectedPath: '/me/planner/tasks?$filter=percentComplete ne 100' },
  { name: 'get-planner-plan', params: { plannerPlanId: 'p1' }, expectedPath: '/planner/plans/p1' },
  { name: 'list-plan-tasks', params: { plannerPlanId: 'p1' }, expectedPath: '/planner/plans/p1/tasks' },
  { name: 'get-planner-task', params: { plannerTaskId: 't1' }, expectedPath: '/planner/tasks/t1' },
  { name: 'get-planner-task-details', params: { plannerTaskId: 't1' }, expectedPath: '/planner/tasks/t1/details' },
  { name: 'list-plan-buckets', params: { plannerPlanId: 'p1' }, expectedPath: '/planner/plans/p1/buckets' },
  { name: 'get-planner-bucket', params: { plannerBucketId: 'b1' }, expectedPath: '/planner/buckets/b1' },
  { name: 'list-mail-messages', params: {}, expectedPath: '/me/messages' },
  { name: 'list-mail-folders', params: {}, expectedPath: '/me/mailFolders' },
  { name: 'list-mail-child-folders', params: { mailFolderId: 'f1' }, expectedPath: '/me/mailFolders/f1/childFolders' },
  { name: 'list-mail-folder-messages', params: { mailFolderId: 'f1' }, expectedPath: '/me/mailFolders/f1/messages' },
  { name: 'get-mail-message', params: { messageId: 'm1' }, expectedPath: '/me/messages/m1' },
  { name: 'list-mail-attachments', params: { messageId: 'm1' }, expectedPath: '/me/messages/m1/attachments' },
  { name: 'get-mail-attachment', params: { messageId: 'm1', attachmentId: 'a1' }, expectedPath: '/me/messages/m1/attachments/a1' },
  { name: 'list-mail-rules', params: { mailFolderId: 'f1' }, expectedPath: '/me/mailFolders/f1/messageRules' },
  { name: 'get-mailbox-settings', params: {}, expectedPath: '/me/mailboxSettings' },
  { name: 'search-mail-messages', params: { query: 'invoice' }, expectedPath: '/me/messages?$search="invoice"' },
  { name: 'list-onenote-notebooks', params: {}, expectedPath: '/me/onenote/notebooks' },
  { name: 'list-onenote-notebook-sections', params: { notebookId: 'n1' }, expectedPath: '/me/onenote/notebooks/n1/sections' },
  { name: 'list-all-onenote-sections', params: {}, expectedPath: '/me/onenote/sections' },
  { name: 'list-onenote-section-pages', params: { onenoteSectionId: 's1' }, expectedPath: '/me/onenote/sections/s1/pages' },
  { name: 'get-onenote-page-content', params: { onenotePageId: 'p1' }, expectedPath: '/me/onenote/pages/p1/content' },
  { name: 'get-onenote-page-as-markdown', params: { onenotePageId: 'p1' }, expectedPath: '/me/onenote/pages/p1/content' },
  { name: 'search-onenote-pages', params: { titleSubstring: 'meeting' }, expectedPath: "/me/onenote/pages?$filter=contains(title,'meeting')" },
  { name: 'get-current-user', params: {}, expectedPath: '/me' },
  { name: 'get-my-profile-photo', params: {}, expectedPath: '/me/photo/$value' },
  { name: 'list-calendar-events', params: {}, expectedPath: '/me/events' },
  { name: 'get-calendar-event', params: { eventId: 'e1' }, expectedPath: '/me/events/e1' },
  { name: 'list-specific-calendar-events', params: { calendarId: 'c1' }, expectedPath: '/me/calendars/c1/events' },
  { name: 'get-specific-calendar-event', params: { calendarId: 'c1', eventId: 'e1' }, expectedPath: '/me/calendars/c1/events/e1' },
  {
    name: 'get-calendar-view',
    params: { startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' },
    expectedPath: '/me/calendarView?startDateTime=2026-04-01T00:00:00Z&endDateTime=2026-05-01T00:00:00Z',
  },
  {
    name: 'get-specific-calendar-view',
    params: { calendarId: 'c1', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' },
    expectedPath: '/me/calendars/c1/calendarView?startDateTime=2026-04-01T00:00:00Z&endDateTime=2026-05-01T00:00:00Z',
  },
  {
    name: 'list-calendar-event-instances',
    params: { calendarId: 'c1', eventId: 'e1', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' },
    expectedPath: '/me/calendars/c1/events/e1/instances?startDateTime=2026-04-01T00:00:00Z&endDateTime=2026-05-01T00:00:00Z',
  },
  { name: 'list-calendars', params: {}, expectedPath: '/me/calendars' },
  { name: 'list-calendar-events-delta', params: {}, expectedPath: '/me/events/delta()' },
  {
    name: 'list-calendar-view-delta',
    params: { startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' },
    expectedPath: '/me/calendarView/delta()?startDateTime=2026-04-01T00:00:00Z&endDateTime=2026-05-01T00:00:00Z',
  },
  { name: 'list-chat-members', params: { chatId: 'ch1' }, expectedPath: '/chats/ch1/members' },
  { name: 'list-joined-teams', params: {}, expectedPath: '/me/joinedTeams' },
  { name: 'get-team', params: { teamId: 'tm1' }, expectedPath: '/teams/tm1' },
  { name: 'list-team-channels', params: { teamId: 'tm1' }, expectedPath: '/teams/tm1/channels' },
  { name: 'get-team-channel', params: { teamId: 'tm1', channelId: 'ch1' }, expectedPath: '/teams/tm1/channels/ch1' },
  {
    name: 'next-page',
    params: { url: 'https://graph.microsoft.com/v1.0/me/messages?$skip=10' },
    expectedPath: '/me/messages?$skip=10',
  },
];

describe('all commands build correct Graph URL', () => {
  it.each(pathFixtures)('$name calls $expectedPath', async ({ name, params, expectedPath }) => {
    const url = await capturedUrl(name, params);
    expect(url).toBe(`https://graph.microsoft.com/v1.0${expectedPath}`);
  });
});
