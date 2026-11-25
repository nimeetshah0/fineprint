import fs from 'node:fs';
import { extractText, getDocumentProxy } from 'unpdf';

export interface ParsedPdf {
  fileName: string;
  pageCount: number;
  paragraphs: Array<{ index: number; page: number; text: string }>;
}

export async function parsePdfBuffer(
  input: Buffer | Uint8Array,
  fileName = 'contract.pdf'
): Promise<ParsedPdf> {
  // pdfjs rejects Node.js Buffer even though Buffer instanceof Uint8Array is true
  const data = new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  const pdf = await getDocumentProxy(data);
  const { totalPages, text } = await extractText(pdf, { mergePages: false });

  const pages = Array.isArray(text) ? text : [String(text || '')];
  const paragraphs: Array<{ index: number; page: number; text: string }> = [];
  let idx = 0;

  pages.forEach((pageText, pIdx) => {
    const pageNum = pIdx + 1;
    const rawBlocks = pageText
      .replace(/\r\n/g, '\n')
      .split(/\n{2,}/)
      .map((b) => b.replace(/\s+/g, ' ').trim())
      .filter((b) => b.length > 0);

    for (const block of rawBlocks) {
      paragraphs.push({ index: idx++, page: pageNum, text: block });
    }
  });

  return { fileName, pageCount: totalPages, paragraphs };
}

/**
 * Extracts clean text paragraphs from a local PDF purely in memory (zero disk retention).
 */
export async function parsePdfFile(filePath: string): Promise<ParsedPdf> {
  const buffer = fs.readFileSync(filePath);
  const fileName = filePath.split(/[\\/]/).pop() || 'contract.pdf';
  return parsePdfBuffer(buffer, fileName);
}
