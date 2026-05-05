import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { convertToMarkdown } from './markdown-pipeline.ts';
import { isPlainTextFilename } from './text-passthrough.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: parsed.error.message });
  const { driveId, itemId } = parsed.data;

  // Same pre-fetch pattern as the PDF variant — Graph `?format=html`
  // only accepts the Office source formats. For plain text / markdown
  // / HTML / JSON we short-circuit and return the raw bytes; turning
  // those through turndown would either round-trip or corrupt them.
  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const name = (meta.value as { name?: string }).name ?? '';

  if (isPlainTextFilename(name)) {
    return graph.getBinary(`/drives/${driveId}/items/${itemId}/content`);
  }
  return convertToMarkdown(graph, `/drives/${driveId}/items/${itemId}/content?format=html`);
};

const meta: CommandMeta = {
  summary:
    'Download a OneDrive / SharePoint file converted to markdown. Graph converts the source to HTML on the fly (`?format=html`), then this CLI runs turndown over it locally to produce clean markdown. Source must be one of the Office formats Graph supports. For plain-text source extensions the command short-circuits and returns raw bytes instead. Worst-case wall-clock is roughly two 60s round-trips back-to-back.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content?format=html',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description:
        'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, ' +
        'or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.',
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the file to convert. Returned by `list-folder-files` or `search-onedrive-files`.' },
  ],
  example: "ask-marcel download-drive-item-as-markdown --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ contentType: "text/markdown", size: <chars>, text: "..." }` for the converted case, or `{ "@microsoft.graph.downloadUrl": "..." }` / `{ contentType, size, base64 }` raw envelope for plain-text source extensions.',
};

export { execute, meta, schema };
