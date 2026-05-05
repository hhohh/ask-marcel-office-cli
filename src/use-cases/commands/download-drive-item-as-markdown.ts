import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { convertToMarkdown } from './markdown-pipeline.ts';
import { isPlainTextFilename } from './text-passthrough.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
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
    'Download a OneDrive / SharePoint file converted to markdown. **Currently broken upstream:** as of 2026-05 Microsoft Graph `?format=html` returns `Sandbox_InputFormatNotSupported` for every Office input format we have tested (docx, pptx, xlsx); the CLI surfaces that error verbatim so an agent can detect it. Use `download-drive-item-as-pdf` until Microsoft restores HTML conversion. Plain-text source extensions still short-circuit and return raw bytes locally — those keep working.',
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
