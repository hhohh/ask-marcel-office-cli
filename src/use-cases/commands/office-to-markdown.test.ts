import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { buildSampleDocx, buildSampleXlsx } from '../../test-helpers/office-fixtures.ts';
import { officeToMarkdown } from './office-to-markdown.ts';

const noopGraph = (overrides: Partial<GraphClient>): GraphClient => ({
  get: async () => ok({}),
  post: async () => ok({}),
  getBinary: async () => ok({}),
  fetchUrl: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
  ...overrides,
});

const toBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

describe('officeToMarkdown — extension dispatch', () => {
  it('returns the raw getBinary envelope for plain-text source extensions (txt/md/json/etc.) without converting', async () => {
    const graph = noopGraph({
      getBinary: async (path) => {
        expect(path).toBe('/drives/d1/items/i1/content');
        return ok({ contentType: 'text/plain', size: 5, text: 'hello' });
      },
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'notes.txt');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect((result.value as { text: string }).text).toBe('hello');
    }
  });

  it('routes docx through mammoth → turndown via the docx-to-markdown helper', async () => {
    const docxBytes = await buildSampleDocx();
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'application/octet-stream', size: docxBytes.byteLength, base64: toBase64(docxBytes) }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'report.docx');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const env = result.value as { contentType: string; text: string };
      expect(env.contentType).toBe('text/markdown');
      expect(env.text).toContain('# Sample Heading');
    }
  });

  it('routes xlsx through sheetjs → markdown table per sheet via the xlsx-to-markdown helper', async () => {
    const xlsxBytes = buildSampleXlsx();
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'application/octet-stream', size: xlsxBytes.byteLength, base64: toBase64(xlsxBytes) }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'data.xlsx');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const env = result.value as { contentType: string; text: string };
      expect(env.contentType).toBe('text/markdown');
      expect(env.text).toContain('## Sheet1');
      expect(env.text).toContain('## Sheet2');
    }
  });

  it('routes loop/fluid/wbtx/whiteboard through the existing Graph ?format=html pipeline (the four inputs Microsoft documents)', async () => {
    let calledPath = '';
    const graph = noopGraph({
      getBinary: async (path) => {
        calledPath = path;
        return ok({ contentType: 'text/html', size: 13, text: '<p>loop x</p>' });
      },
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'planning.loop');
    expect(result.ok).toBe(true);
    expect(calledPath).toBe('/drives/d1/items/i1/content?format=html');
  });

  it('errs with a clear pptx-specific hint pointing at the *-as-pdf sibling for pptx', async () => {
    const graph = noopGraph({});
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'deck.pptx');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(415);
      expect(result.error.message).toContain('pptx not supported');
      expect(result.error.message).toContain('*-as-pdf');
      expect(result.error.message).toContain('vision-capable LLM');
    }
  });

  it('errs with the generic 38-input-set hint for every other extension (pdf, rtf, odt, etc.)', async () => {
    const graph = noopGraph({});
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'report.pdf');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(415);
      expect(result.error.message).toContain('pdf not supported');
      expect(result.error.message).toContain('38 input extensions');
    }
  });

  it('errs with `<no-extension>` placeholder when the filename has no dot', async () => {
    const graph = noopGraph({});
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'README');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('<no-extension>');
    }
  });

  it('errs with `<no-extension>` placeholder when the filename ends with a dot but no extension', async () => {
    const graph = noopGraph({});
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'oddfile.');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('<no-extension>');
    }
  });

  it('propagates getBinary api_error from the docx path', async () => {
    const graph = noopGraph({
      getBinary: async () => ({ ok: false, error: { type: 'api_error' as const, status: 404, message: 'not found' } }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'report.docx');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
    }
  });

  it('propagates getBinary api_error from the xlsx path', async () => {
    const graph = noopGraph({
      getBinary: async () => ({ ok: false, error: { type: 'api_error' as const, status: 401, message: 'unauthorized' } }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'data.xlsx');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(401);
    }
  });

  it('errs cleanly when getBinary returns an envelope with neither base64 nor text (defensive guard for unknown shapes)', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ surprise: true }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'report.docx');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('unexpected getBinary envelope');
    }
  });

  it('reads the text envelope path of getBinary for xlsx (rare but valid for very small workbooks)', async () => {
    // synthetic: xlsx text envelope cannot really happen, but the decoder must still cover the text branch.
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'text/plain', size: 4, text: 'PK' }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'data.xlsx');
    // sheetjs will reject; we only need to verify the text-decoder branch was exercised.
    expect(result.ok).toBe(false);
  });
});
