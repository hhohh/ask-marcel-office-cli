import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { convertToMarkdown } from './markdown-pipeline.ts';
import { isPlainTextFilename } from './text-passthrough.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  versionId: z.string().min(1),
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId, versionId } = parsed.data;

  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const name = (meta.value as { name?: string }).name ?? '';

  if (isPlainTextFilename(name)) {
    return graph.getBinary(`/drives/${driveId}/items/${itemId}/versions/${versionId}/content`);
  }
  return convertToMarkdown(graph, `/drives/${driveId}/items/${itemId}/versions/${versionId}/content?format=html`);
};

const meta: CommandMeta = {
  summary:
    'Download a *historical version* of a OneDrive / SharePoint file converted to markdown. **Currently broken upstream:** as of 2026-05 the historical-version `?format=html` endpoint returns `Forbidden` (the current-version sibling returns `Sandbox_InputFormatNotSupported`). Both failure modes trace back to Microsoft disabling HTML conversion at the Office Online sandbox. Use `download-drive-item-version-as-pdf` until Microsoft restores it. Plain-text source extensions still short-circuit to raw bytes.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/versions/{version-id}/content?format=html',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitemversion-get-content',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description:
        'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, ' +
        'or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.',
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the file. Returned by `list-folder-files` or `search-onedrive-files`.' },
    {
      name: 'version-id',
      key: 'versionId',
      required: true,
      description:
        'driveItemVersion ID. Returned by `ask-marcel list-drive-item-versions`. ' +
        'Pick a non-current version (the first entry is the live file and Graph rejects this endpoint for it).',
    },
  ],
  example: "ask-marcel download-drive-item-version-as-markdown --drive-id 'b!1234' --item-id '01ABC' --version-id '4.0'",
  responseShape: '`{ contentType: "text/markdown", size: <chars>, text: "..." }` for the converted case; raw-bytes envelope for plain-text source extensions.',
};

export { execute, meta, schema };
