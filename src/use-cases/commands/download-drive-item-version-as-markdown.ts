import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { officeToMarkdown } from './office-to-markdown.ts';

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

  return officeToMarkdown(graph, `/drives/${driveId}/items/${itemId}/versions/${versionId}/content`, name);
};

const meta: CommandMeta = {
  summary:
    'Download a *historical version* of a OneDrive / SharePoint file converted to markdown. Same local conversion pipeline as `download-drive-item-as-markdown`: docx via mammoth, xlsx via sheetjs (markdown tables per sheet), csv as a markdown table, plus plain-text passthrough (txt/md/html/json/yaml/etc.). For pptx use `download-drive-item-version-as-pdf` (PDF preserves slide layout). For pdf/rtf/odt/etc. also use the PDF sibling. Loop/Fluid/Whiteboard use Graph `?format=html` (the four inputs Microsoft documents — https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format).',
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
