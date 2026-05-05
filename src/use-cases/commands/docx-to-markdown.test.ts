import { describe, expect, it } from 'bun:test';
import { buildMalformedDocx, buildSampleDocx } from '../../test-helpers/office-fixtures.ts';
import { docxToMarkdown, promoteFirstRowToThead } from './docx-to-markdown.ts';

describe('docxToMarkdown', () => {
  it('converts a docx into a markdown envelope with heading, bold/italic, table, and inline image as data URI', async () => {
    const bytes = await buildSampleDocx();
    const result = await docxToMarkdown(bytes);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.contentType).toBe('text/markdown');
      expect(result.value.size).toBeGreaterThan(0);
      expect(result.value.text).toContain('# Sample Heading');
      expect(result.value.text).toContain('**world**');
      expect(result.value.text).toContain('_italic_');
      expect(result.value.text).toMatch(/\|\s*A\s*\|\s*B\s*\|/);
      expect(result.value.text).toContain('data:image/png;base64,');
    }
  });

  it('propagates the api_error from the mammoth adapter when the input is not a valid docx', async () => {
    const result = await docxToMarkdown(buildMalformedDocx());
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('docx conversion failed');
    }
  });
});

describe('promoteFirstRowToThead — table walker edge cases', () => {
  it('returns HTML unchanged when there is no `<table>` at all', () => {
    const input = '<h1>title</h1><p>just text, no tables</p>';
    expect(promoteFirstRowToThead(input)).toBe(input);
  });

  it('returns HTML unchanged when a `<table>` opens but never closes (malformed)', () => {
    const input = '<p>before</p><table><tr><td>A</td></tr><p>truncated';
    expect(promoteFirstRowToThead(input)).toBe(input);
  });

  it('promotes the first row when the table is wrapped in `<tbody>` (defensive — mammoth does not currently emit this, but the walker handles it)', () => {
    const input = '<table><tbody><tr><td>H</td></tr><tr><td>1</td></tr></tbody></table>';
    expect(promoteFirstRowToThead(input)).toBe('<table><thead><tr><td>H</td></tr></thead><tbody><tr><td>1</td></tr></tbody></table>');
  });

  it('returns the original table when there is no `</tr>` inside (degenerate empty table)', () => {
    const input = '<table></table>';
    expect(promoteFirstRowToThead(input)).toBe('<table></table>');
  });

  it('walks past content between two tables and promotes both', () => {
    const input = '<table><tr><td>A</td></tr></table>middle<table><tr><td>B</td></tr></table>';
    const out = promoteFirstRowToThead(input);
    expect(out).toContain('<thead><tr><td>A</td></tr></thead>');
    expect(out).toContain('middle');
    expect(out).toContain('<thead><tr><td>B</td></tr></thead>');
  });
});
