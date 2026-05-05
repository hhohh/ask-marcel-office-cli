import { Document, HeadingLevel, ImageRun, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from 'docx';
import * as XLSX from 'xlsx';

const TINY_PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==';

const tinyPngBytes = (): Uint8Array => {
  const binary = atob(TINY_PNG_BASE64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
};

const buildSampleDocx = async (): Promise<Uint8Array> => {
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Sample Heading')] }),
          new Paragraph({
            children: [
              new TextRun('Hello '),
              new TextRun({ text: 'world', bold: true }),
              new TextRun(', this is '),
              new TextRun({ text: 'italic', italics: true }),
              new TextRun(' text.'),
            ],
          }),
          new Paragraph({
            children: [
              new ImageRun({
                type: 'png',
                data: tinyPngBytes(),
                transformation: { width: 100, height: 100 },
              }),
            ],
          }),
          new Table({
            rows: [
              new TableRow({
                children: [new TableCell({ children: [new Paragraph('A')] }), new TableCell({ children: [new Paragraph('B')] })],
              }),
              new TableRow({
                children: [new TableCell({ children: [new Paragraph('1')] }), new TableCell({ children: [new Paragraph('2')] })],
              }),
            ],
          }),
        ],
      },
    ],
  });
  const buffer = await Packer.toBuffer(doc);
  return new Uint8Array(buffer);
};

const buildSampleXlsx = (): Uint8Array => {
  const sheet1 = XLSX.utils.aoa_to_sheet([
    ['Name', 'Age', 'City'],
    ['Alice', 30, 'Paris'],
    ['Bob', 25, 'Berlin'],
  ]);
  const sheet2 = XLSX.utils.aoa_to_sheet([
    ['Product', 'Price'],
    ['Widget', 9.99],
    ['Gadget', 14.5],
  ]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet1, 'Sheet1');
  XLSX.utils.book_append_sheet(workbook, sheet2, 'Sheet2');
  const arrayBuffer = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
  return new Uint8Array(arrayBuffer);
};

const buildMalformedDocx = (): Uint8Array => new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00, 0x00, 0x00]);

const buildMalformedXlsx = (): Uint8Array => new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0xff, 0xff, 0xff, 0xff]);

export { buildMalformedDocx, buildMalformedXlsx, buildSampleDocx, buildSampleXlsx };
