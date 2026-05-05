import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { isPlainTextFilename } from './text-passthrough.ts';

const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  versionId: z.string().min(1),
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: parsed.error.message });
  const { driveId, itemId, versionId } = parsed.data;

  // Pre-fetch the driveItem for its filename. Same pre-check as the
  // non-versioned variant: short-circuit to a raw-bytes download for
  // plain-text source extensions instead of letting Graph reject the
  // conversion call with a confusing 4xx.
  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const name = (meta.value as { name?: string }).name ?? '';

  if (isPlainTextFilename(name)) {
    return graph.getBinary(`/drives/${driveId}/items/${itemId}/versions/${versionId}/content`);
  }
  return graph.getBinary(`/drives/${driveId}/items/${itemId}/versions/${versionId}/content?format=pdf`);
};

const meta: CommandMeta = {
  summary:
    'Download a *historical version* of a OneDrive / SharePoint file converted to PDF on the fly by Graph. Same shape as `download-drive-item-as-pdf` plus a `--version-id`. Graph refuses to serve the *current* version through this endpoint — for the current version use `download-drive-item-as-pdf` instead. Plain-text source extensions short-circuit to a raw-bytes download.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/versions/{version-id}/content?format=pdf',
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
        'Pick a non-current version — the first entry (e.g. `12.0`) is the live file and Graph rejects this endpoint for it; use `value[1]` or older.',
    },
  ],
  example: "ask-marcel download-drive-item-version-as-pdf --drive-id 'b!1234' --item-id '01ABC' --version-id '4.0'",
  responseShape:
    '`{ "@microsoft.graph.downloadUrl": "..." }` for the typical 302 case, or `{ contentType, size, base64 }` when Graph streams bytes directly. Raw-bytes envelope for plain-text source extensions.',
};

export { execute, meta, schema };
