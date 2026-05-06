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

/**
 * Microsoft's SharePoint streamContent endpoint (which Graph redirects
 * to for `/drives/{}/items/{}/versions/{ver}/content`) rejects the
 * Teams web client token with `403 logicalPermissionAccessDenied`. The
 * URL-returning siblings (`*-version-content`, `*-version-as-pdf`) hide
 * this because they don't follow the redirect — they return a URL
 * which 403s when fetched. The markdown command can't hide it: it
 * actually has to fetch the bytes.
 *
 * When we see this specific error, rewrite to actionable guidance.
 */
const isLogicalPermissionDenied = (e: GraphError): boolean => e.type === 'api_error' && e.status === 403 && e.message.includes('logicalPermissionAccessDenied');

const augmentVersionAccessDenied = (e: GraphError): GraphError => {
  if (!isLogicalPermissionDenied(e)) return e;
  return {
    type: 'api_error',
    status: 403,
    message:
      'historical-version stream content is blocked by Microsoft for the Teams web client token this CLI uses (innerError code logicalPermissionAccessDenied — see https://aka.ms/ODSPS2SAuthOnboarding). Workarounds: use `download-drive-item-as-markdown` (without --version-id) for the *current* version content, or run from an environment with elevated ODSP scopes. The URL-returning siblings (`download-drive-item-version-content`, `download-drive-item-version-as-pdf`) appear to succeed but return URLs that 403 the same way when followed.',
  };
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId, versionId } = parsed.data;

  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const name = (meta.value as { name?: string }).name ?? '';

  const result = await officeToMarkdown(graph, `/drives/${driveId}/items/${itemId}/versions/${versionId}/content`, name);
  if (result.ok) return result;
  return err(augmentVersionAccessDenied(result.error));
};

const meta: CommandMeta = {
  summary:
    'Download a *historical version* of a OneDrive / SharePoint file converted to markdown. **Known Teams-token limit:** Microsoft’s SharePoint streamContent endpoint blocks the Teams web client token with `403 logicalPermissionAccessDenied` for historical-version bytes — the command can fetch the bytes for the *current* version (no --version-id) but not for older versions. Same local conversion pipeline as `download-drive-item-as-markdown`: docx via mammoth, xlsx via sheetjs (markdown tables per sheet), csv as a markdown table, plus plain-text passthrough. For pptx use `download-drive-item-version-as-pdf` (which returns a URL — though that URL hits the same Teams-token wall when followed). Loop/Fluid/Whiteboard use Graph `?format=html` (the four inputs Microsoft documents).',
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
