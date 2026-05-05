import { describe, expect, it } from 'bun:test';
import { buildMalformedXlsx, buildSampleXlsx } from '../test-helpers/office-fixtures.ts';
import { readSheetsAsCsv } from './sheetjs-adapter.ts';

describe('readSheetsAsCsv', () => {
  it('returns one entry per worksheet, each with the sheet name and CSV-rendered cells', () => {
    const result = readSheetsAsCsv(buildSampleXlsx());
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toHaveLength(2);
      expect(result.value[0]?.name).toBe('Sheet1');
      expect(result.value[0]?.csv).toContain('Name,Age,City');
      expect(result.value[0]?.csv).toContain('Alice,30,Paris');
      expect(result.value[1]?.name).toBe('Sheet2');
      expect(result.value[1]?.csv).toContain('Product,Price');
      expect(result.value[1]?.csv).toContain('Widget,9.99');
    }
  });

  it('returns err({ type: api_error }) when the bytes are not a valid xlsx archive', () => {
    const result = readSheetsAsCsv(buildMalformedXlsx());
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(500);
      expect(result.error.message).toContain('xlsx parse failed');
    }
  });
});
