import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { isPlainTextFilename } from './text-passthrough.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) throw new Error(`validation failed: ${parsed.error.message}`);
  const { driveId, itemId } = parsed.data;

  // Pre-fetch the driveItem metadata to read its filename. Graph's
  // ?format=pdf only accepts Office source formats — for plain-text
  // and other unsupported extensions we short-circuit and return raw
  // bytes instead of letting Graph reject the call with a confusing 4xx.
  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const name = (meta.value as { name?: string }).name ?? '';

  if (isPlainTextFilename(name)) {
    return graph.getBinary(`/drives/${driveId}/items/${itemId}/content`);
  }
  return graph.getBinary(`/drives/${driveId}/items/${itemId}/content?format=pdf`);
};

const meta: CommandMeta = {
  summary:
    'Download a OneDrive / SharePoint file converted to PDF on the fly by Graph (`?format=pdf`). Source must be one of the Office formats Graph supports — doc, docx, ppt, pptx, xls, xlsx, rtf, csv, odp, ods, odt, etc. The command pre-fetches the filename and short-circuits to a raw download for plain-text source extensions (txt, md, html, json, …) since Graph would reject those anyway. Worst-case wall-clock is two 60s round-trips back-to-back.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content?format=pdf',
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
  example: "ask-marcel download-drive-item-as-pdf --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ "@microsoft.graph.downloadUrl": "..." }` for the typical 302 case, or `{ contentType, size, base64 }` when Graph streams bytes directly. For unsupported source extensions, returns the raw file bytes (same envelope) without the PDF conversion step.',
};

export { execute, meta, schema };
