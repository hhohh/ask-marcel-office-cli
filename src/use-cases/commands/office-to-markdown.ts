import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { docxToMarkdown } from './docx-to-markdown.ts';
import { convertToMarkdown } from './markdown-pipeline.ts';
import { isPlainTextFilename } from './text-passthrough.ts';
import { xlsxToMarkdown } from './xlsx-to-markdown.ts';

/**
 * Strategy dispatcher for `*-as-markdown` commands. Picks the right
 * conversion path based on the file extension and routes through it:
 *
 * - plain-text   → raw bytes envelope (Graph getBinary, no conversion)
 * - docx         → mammoth → turndown
 * - xlsx         → sheetjs → markdown table per sheet
 * - loop / fluid / wbtx / whiteboard → existing Graph ?format=html
 *                                      pipeline (the only inputs Graph
 *                                      HTML conversion actually accepts)
 * - everything else (pptx, pdf, rtf, odt, …) → err pointing at the
 *                                              corresponding *-as-pdf
 *                                              command (which accepts
 *                                              38 input extensions)
 */

const HTML_FORMAT_INPUTS: ReadonlySet<string> = new Set(['loop', 'fluid', 'wbtx', 'whiteboard']);

const PPTX_HINT =
  'pptx not supported by `*-as-markdown`. Use the corresponding `*-as-pdf` command — Graph PDF conversion preserves slide layout, and a vision-capable LLM reads it more reliably than flattened slide-by-slide bullets.';

const GENERIC_HINT = (ext: string): string =>
  `${ext} not supported by \`*-as-markdown\`. Use the corresponding \`*-as-pdf\` command — Graph \`?format=pdf\` accepts 38 input extensions including this one.`;

const extensionOf = (filename: string): string => {
  const dot = filename.lastIndexOf('.');
  if (dot === -1 || dot === filename.length - 1) return '';
  return filename.slice(dot + 1).toLowerCase();
};

const officeToMarkdown = async (graph: GraphClient, contentPath: string, filename: string): Promise<Result<unknown, GraphError>> => {
  if (isPlainTextFilename(filename)) {
    return graph.getBinary(contentPath);
  }
  const ext = extensionOf(filename);
  if (ext === 'docx') {
    const bytes = await graph.getBinary(contentPath);
    if (!bytes.ok) return bytes;
    const blob = bytes.value as { base64?: string; text?: string };
    const raw = decodeRawBytes(blob);
    if (!raw.ok) return raw;
    return docxToMarkdown(raw.value);
  }
  if (ext === 'xlsx') {
    const bytes = await graph.getBinary(contentPath);
    if (!bytes.ok) return bytes;
    const blob = bytes.value as { base64?: string; text?: string };
    const raw = decodeRawBytes(blob);
    if (!raw.ok) return raw;
    return xlsxToMarkdown(raw.value);
  }
  if (HTML_FORMAT_INPUTS.has(ext)) {
    return convertToMarkdown(graph, `${contentPath}?format=html`);
  }
  if (ext === 'pptx') {
    return err({ type: 'api_error', status: 415, message: PPTX_HINT });
  }
  return err({ type: 'api_error', status: 415, message: GENERIC_HINT(ext === '' ? '<no-extension>' : ext) });
};

const decodeRawBytes = (blob: { base64?: string; text?: string }): Result<Uint8Array, GraphError> => {
  if (typeof blob.base64 === 'string') {
    const binary = atob(blob.base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return { ok: true, value: bytes };
  }
  if (typeof blob.text === 'string') {
    return { ok: true, value: new TextEncoder().encode(blob.text) };
  }
  return err({ type: 'api_error', status: 500, message: 'unexpected getBinary envelope: missing both base64 and text fields' });
};

export { officeToMarkdown };
