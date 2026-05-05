import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { buildShareToken, extractSharepointUrls } from './sharepoint-link-extractor.ts';

const MAX_LINKS = 25; // Hardening #4: cap fan-out

type ResolvedLink = {
  readonly url: string;
  readonly driveId?: string;
  readonly itemId?: string;
  readonly name?: string;
  readonly webUrl?: string;
  readonly error?: string;
};

type LinkExtractionSummary = {
  readonly messageId: string;
  readonly subject?: string;
  readonly links: ReadonlyArray<ResolvedLink>;
  readonly truncated: boolean;
  readonly skippedCount: number;
};

const schema = z.object({ messageId: z.string().min(1) });

const resolveOne = async (graph: GraphClient, url: string): Promise<ResolvedLink> => {
  const token = buildShareToken(url);
  const result = await graph.get(`/shares/${token}/driveItem`);
  if (!result.ok) return { url, error: result.error.type === 'api_error' ? result.error.message : `${result.error.type}: ${result.error.message}` };
  const item = result.value as { id?: string; name?: string; webUrl?: string; parentReference?: { driveId?: string } };
  return {
    url,
    driveId: item.parentReference?.driveId,
    itemId: item.id,
    name: item.name,
    webUrl: item.webUrl,
  };
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<LinkExtractionSummary, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) throw new Error(`validation failed: ${parsed.error.message}`);
  const { messageId } = parsed.data;

  // 1. Pull the mail subject + body.
  const message = await graph.get(`/me/messages/${messageId}?$select=subject,body`);
  if (!message.ok) return message;
  const m = message.value as { subject?: string; body?: { content?: string } };
  const body = m.body?.content ?? '';

  // 2. Find every *.sharepoint.com URL.
  const allUrls = extractSharepointUrls(body);
  const truncated = allUrls.length > MAX_LINKS;
  const skippedCount = truncated ? allUrls.length - MAX_LINKS : 0;
  const kept = truncated ? allUrls.slice(0, MAX_LINKS) : allUrls;

  // 3. Resolve each via /shares/{token}/driveItem.
  const links = await Promise.all(kept.map((u) => resolveOne(graph, u)));

  return ok({
    messageId,
    subject: m.subject,
    links,
    truncated,
    skippedCount,
  });
};

const meta: CommandMeta = {
  summary:
    'Find every `*.sharepoint.com` URL in the body of a single Outlook email and resolve each one to its driveItem (driveId, itemId, name, webUrl) so the agent can feed those into `download-drive-item-as-pdf` / `-as-markdown` etc. Read-only — no conversion happens here. Capped at 25 unique URLs per call to bound fan-out (returns `truncated: true` and `skippedCount` when the body has more); duplicate URLs are deduplicated. Per-link errors are captured inside each entry instead of failing the whole call.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-get',
  options: [
    {
      name: 'message-id',
      key: 'messageId',
      required: true,
      description: 'Outlook message ID. Returned by `list-mail-messages` or `list-mail-folder-messages`.',
    },
  ],
  example: "ask-marcel extract-sharepoint-links-in-mail --message-id 'AAMkADk0...'",
  responseShape:
    '`{ messageId, subject, links: [{ url, driveId, itemId, name, webUrl } | { url, error }], truncated, skippedCount }` — one entry per unique SharePoint URL found in the body, ordered by first occurrence.',
};

export { execute, meta, schema };
export type { LinkExtractionSummary, ResolvedLink };
