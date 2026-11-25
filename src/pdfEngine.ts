import fs from 'node:fs';
import { extractText, getDocumentProxy } from 'unpdf';

export interface ParsedPdf {
  fileName: string;
  pageCount: number;
  paragraphs: Array<{ index: number; page: number; text: string }>;
}

/**
 * Extracts clean text paragraphs from a PDF purely in memory (zero disk retention).
 * Works in Node.js and edge runtimes (Cloudflare Workers, etc.).
 */
export async function parsePdfFile(filePath: string): Promise<ParsedPdf> {
  const buffer = fs.readFileSync(filePath);
  const data = new Uint8Array(buffer);
  const pdf = await getDocumentProxy(data);
  const { totalPages, text } = await extractText(pdf, { mergePages: false });

  const pages = Array.isArray(text) ? text : [String(text || '')];
  const paragraphs: Array<{ index: number; page: number; text: string }> = [];
  let idx = 0;

  pages.forEach((pageText, pIdx) => {
    const pageNum = pIdx + 1;
    // Split on double newlines or clear section boundaries while cleaning extra spaces
    const rawBlocks = pageText
      .replace(/\r\n/g, '\n')
      .split(/\n{2,}/)
      .map((b) => b.replace(/\s+/g, ' ').trim())
      .filter((b) => b.length > 0);

    for (const block of rawBlocks) {
      paragraphs.push({ index: idx++, page: pageNum, text: block });
    }
  });

  const fileName = filePath.split(/[\\/]/).pop() || 'contract.pdf';
  return { fileName, pageCount: totalPages, paragraphs };
}
