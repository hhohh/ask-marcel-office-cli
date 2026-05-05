import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { htmlToMarkdown } from './html-to-markdown.ts';
import { embedInlineImages, type InlineAttachment } from './inline-image-embedder.ts';

/**
 * Orchestrate the four steps that turn a Graph `?format=html` content
 * call into a markdown envelope:
 *
 *   1. getBinary(contentPath)
 *   2. if Graph returned a 302 downloadUrl, follow via fetchUrl (host
 *      allow-list enforced inside fetchUrl, Hardening #3)
 *   3. embedInlineImages over any cid: refs (Hardening #1: image/* only)
 *   4. htmlToMarkdown via turndown
 *
 * Returns `{ contentType: 'text/markdown', size, text }` on success.
 */

type MarkdownEnvelope = { readonly contentType: 'text/markdown'; readonly size: number; readonly text: string };

const extractHtml = async (graph: GraphClient, binaryResult: Record<string, unknown>): Promise<Result<string, GraphError>> => {
  if (typeof binaryResult.text === 'string') {
    return ok(binaryResult.text);
  }
  if (typeof binaryResult['@microsoft.graph.downloadUrl'] === 'string') {
    const followed = await graph.fetchUrl(binaryResult['@microsoft.graph.downloadUrl']);
    if (!followed.ok) return followed;
    const v = followed.value as Record<string, unknown>;
    if (typeof v.text === 'string') return ok(v.text);
    return err({ type: 'api_error', status: 500, message: 'unexpected response shape from fetchUrl: missing text field' });
  }
  return err({ type: 'api_error', status: 500, message: 'unexpected response shape from getBinary: no text or downloadUrl' });
};

const convertToMarkdown = async (
  graph: GraphClient,
  contentPath: string,
  inlineAttachments: ReadonlyArray<InlineAttachment> = []
): Promise<Result<MarkdownEnvelope, GraphError>> => {
  const binary = await graph.getBinary(contentPath);
  if (!binary.ok) return binary;
  const html = await extractHtml(graph, binary.value as Record<string, unknown>);
  if (!html.ok) return html;
  const inlined = inlineAttachments.length > 0 ? embedInlineImages(html.value, inlineAttachments) : html.value;
  const md = htmlToMarkdown(inlined);
  return ok({ contentType: 'text/markdown', size: md.length, text: md });
};

export { convertToMarkdown };
export type { MarkdownEnvelope };
