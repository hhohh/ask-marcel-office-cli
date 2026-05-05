import { describe, expect, it } from 'bun:test';
import { buildMalformedXlsx, buildSampleXlsx } from '../../test-helpers/office-fixtures.ts';
import { csvToMarkdownTable, xlsxToMarkdown } from './xlsx-to-markdown.ts';

describe('xlsxToMarkdown', () => {
  it('converts an xlsx into one `## SheetName` section per sheet, each with a markdown table', () => {
    const result = xlsxToMarkdown(buildSampleXlsx());
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.contentType).toBe('text/markdown');
      expect(result.value.size).toBeGreaterThan(0);
      expect(result.value.text).toContain('## Sheet1');
      expect(result.value.text).toMatch(/\|\s*Name\s*\|\s*Age\s*\|\s*City\s*\|/);
      expect(result.value.text).toMatch(/\|\s*Alice\s*\|\s*30\s*\|\s*Paris\s*\|/);
      expect(result.value.text).toContain('## Sheet2');
      expect(result.value.text).toMatch(/\|\s*Product\s*\|\s*Price\s*\|/);
    }
  });

  it('separates the per-sheet sections with blank lines', () => {
    const result = xlsxToMarkdown(buildSampleXlsx());
    if (result.ok) {
      const sheet1Index = result.value.text.indexOf('## Sheet1');
      const sheet2Index = result.value.text.indexOf('## Sheet2');
      expect(sheet1Index).toBeLessThan(sheet2Index);
      expect(result.value.text.slice(sheet1Index, sheet2Index)).toContain('\n\n');
    }
  });

  it('propagates the api_error from the sheetjs adapter when the input is not a valid xlsx', () => {
    const result = xlsxToMarkdown(buildMalformedXlsx());
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('xlsx parse failed');
    }
  });
});

describe('csvToMarkdownTable', () => {
  it('renders a header row + separator + data rows from a basic CSV', () => {
    const md = csvToMarkdownTable('a,b\n1,2\n3,4');
    expect(md).toBe('| a | b |\n| --- | --- |\n| 1 | 2 |\n| 3 | 4 |');
  });

  it('returns an empty string when the CSV has no non-empty lines', () => {
    expect(csvToMarkdownTable('')).toBe('');
    expect(csvToMarkdownTable('\n\n')).toBe('');
  });

  it('handles cells with embedded commas by respecting CSV quoting', () => {
    const md = csvToMarkdownTable('"a,b",c\n"1,2",3');
    expect(md).toContain('| a,b | c |');
    expect(md).toContain('| 1,2 | 3 |');
  });

  it('handles escaped quotes inside a quoted cell ("" → ")', () => {
    const md = csvToMarkdownTable('a,b\n"he said ""hi""",2');
    expect(md).toContain('| he said "hi" | 2 |');
  });

  it('pads short rows with empty cells so every row has the same column count', () => {
    const md = csvToMarkdownTable('a,b,c\n1\n2,3');
    expect(md).toContain('| 1 |  |  |');
    expect(md).toContain('| 2 | 3 |  |');
  });
});
