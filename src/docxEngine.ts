import fs from 'node:fs';
import JSZip from 'jszip';
import type { RedlineClause, WordDiffToken } from './types.js';

/**
 * Word-level LCS diff so both the UI and OOXML Track Changes (<w:del>/<w:ins>)
 * highlight exact word changes instead of replacing whole paragraphs.
 */
export function computeWordDiff(original: string, proposed: string): WordDiffToken[] {
  if (original === proposed) {
    return original ? [{ type: 'equal', text: original }] : [];
  }

  const a = tokenizeWords(original);
  const b = tokenizeWords(proposed);

  if (a.length * b.length > 360_000) {
    return [
      { type: 'del', text: original },
      { type: 'ins', text: proposed },
    ];
  }

  const m = a.length;
  const n = b.length;
  const dp: Uint16Array[] = Array.from({ length: m + 1 }, () => new Uint16Array(n + 1));

  for (let i = m - 1; i >= 0; i--) {
    for (let j = n - 1; j >= 0; j--) {
      if (a[i] === b[j]) dp[i][j] = dp[i + 1][j + 1] + 1;
      else dp[i][j] = dp[i + 1][j] >= dp[i][j + 1] ? dp[i + 1][j] : dp[i][j + 1];
    }
  }

  const raw: WordDiffToken[] = [];
  let i = 0;
  let j = 0;
  while (i < m && j < n) {
    if (a[i] === b[j]) {
      raw.push({ type: 'equal', text: a[i] });
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      raw.push({ type: 'del', text: a[i] });
      i++;
    } else {
      raw.push({ type: 'ins', text: b[j] });
      j++;
    }
  }
  while (i < m) raw.push({ type: 'del', text: a[i++] });
  while (j < n) raw.push({ type: 'ins', text: b[j++] });

  const merged: WordDiffToken[] = [];
  for (const tok of raw) {
    const prev = merged[merged.length - 1];
    if (prev && prev.type === tok.type) prev.text += tok.text;
    else merged.push({ ...tok });
  }
  return merged;
}

function tokenizeWords(s: string): string[] {
  return s.match(/\S+|\s+/g) || [];
}

function xmlEscape(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function xmlUnescape(text: string): string {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

export interface ParsedDocx {
  fileName: string;
  buffer: Buffer;
  paragraphs: Array<{ index: number; text: string }>;
}

/**
 * Depth-aware scanner for <w:p>...</w:p> ranges inside word/document.xml.
 * Safely skips self-closing <w:p/> spacer tags and handles nested <w:p> inside
 * text boxes (<w:txbxContent>) without regex corruption.
 */
function findParagraphRanges(docXml: string): Array<{ start: number; end: number; xml: string }> {
  const ranges: Array<{ start: number; end: number; xml: string }> = [];
  const tagRegex = /<(\/?)w:p(?:\s[^>]*)?(\/?)>/g;
  let depth = 0;
  let startIdx = -1;
  let m: RegExpExecArray | null;

  while ((m = tagRegex.exec(docXml)) !== null) {
    const isClose = m[1] === '/';
    const isSelfClose = m[2] === '/';

    if (isSelfClose) continue;
    if (!isClose) {
      if (depth === 0) startIdx = m.index;
      depth++;
    } else {
      depth--;
      if (depth === 0 && startIdx >= 0) {
        const endIdx = m.index + m[0].length;
        ranges.push({
          start: startIdx,
          end: endIdx,
          xml: docXml.slice(startIdx, endIdx),
        });
        startIdx = -1;
      } else if (depth < 0) {
        depth = 0;
      }
    }
  }
  return ranges;
}

/**
 * Parses a .docx from an in-memory Buffer/Uint8Array (works in Cloudflare Workers and Node).
 * Reads the current accepted view (includes <w:ins> text, excludes <w:delText>).
 */
export async function parseDocxBuffer(
  input: Buffer | Uint8Array,
  fileName = 'contract.docx'
): Promise<ParsedDocx> {
  const buffer = Buffer.isBuffer(input) ? input : Buffer.from(input);
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml')?.async('string');
  if (!docXml) {
    throw new Error('Invalid .docx file: missing word/document.xml');
  }

  const paragraphs: Array<{ index: number; text: string }> = [];
  const ranges = findParagraphRanges(docXml);

  ranges.forEach((r, idx) => {
    const withoutDel = r.xml.replace(/<w:del\b[\s\S]*?<\/w:del>/g, '');
    const textRuns = [...withoutDel.matchAll(/<w:t(?:\s+[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) =>
      xmlUnescape(m[1])
    );
    const text = textRuns.join('').trim();
    if (text.length > 0) {
      paragraphs.push({ index: idx, text });
    }
  });

  return { fileName, buffer, paragraphs };
}

/**
 * Extracts numbered plain-text paragraphs from a local .docx file purely in memory.
 */
export async function parseDocxFile(filePath: string): Promise<ParsedDocx> {
  const buffer = fs.readFileSync(filePath);
  const fileName = filePath.split(/[\\/]/).pop() || 'contract.docx';
  return parseDocxBuffer(buffer, fileName);
}

/**
 * Extracts a Google Docs document ID from a standard docs.google.com URL.
 */
export function extractGoogleDocId(urlOrId: string): string | null {
  const trimmed = urlOrId.trim();
  const match = trimmed.match(/docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/i);
  if (match) return match[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(trimmed)) return trimmed;
  return null;
}

/**
 * Fetches a link-shared Google Doc directly as a real .docx binary in memory.
 * Exported Track Changes (.docx) re-upload into Google Docs as native Suggestions!
 */
export async function fetchGoogleDocAsDocx(urlOrId: string): Promise<ParsedDocx> {
  const docId = extractGoogleDocId(urlOrId);
  if (!docId) {
    throw new Error(
      'Invalid Google Docs URL. Expected https://docs.google.com/document/d/<DOC_ID>/edit'
    );
  }
  const exportUrl = `https://docs.google.com/document/d/${docId}/export?format=docx`;
  const res = await fetch(exportUrl, { redirect: 'follow' });
  const contentType = res.headers.get('content-type') || '';

  if (!res.ok || contentType.includes('text/html')) {
    throw new Error(
      'This Google Doc is private or requires sign-in. Either temporarily set sharing to "Anyone with the link can view", or in Google Docs click File → Download → Microsoft Word (.docx) and drop the file into Fineprint.'
    );
  }

  const arrBuf = await res.arrayBuffer();
  const cd = res.headers.get('content-disposition') || '';
  const nameMatch = cd.match(/filename\*?=(?:UTF-8'')?"?([^";\n]+)"?/i);
  const fileName = nameMatch ? decodeURIComponent(nameMatch[1]) : `GoogleDoc_${docId.slice(0, 8)}.docx`;

  return parseDocxBuffer(new Uint8Array(arrBuf), fileName);
}

/**
 * Finds the maximum numeric w:id across one or more OOXML documents
 * so newly emitted revision & comment IDs never collide with existing ones.
 */
function findMaxXmlId(...xmlDocs: Array<string | undefined>): number {
  let maxId = 100;
  for (const xml of xmlDocs) {
    if (!xml) continue;
    for (const m of xml.matchAll(/\bw:id="(\d+)"/g)) {
      const n = Number(m[1]);
      if (Number.isFinite(n) && n > maxId) maxId = n;
    }
  }
  return maxId;
}

/**
 * Extracts the dominant body run properties (<w:rPr>) from a paragraph so patched
 * text keeps the document's font family, size, and color (prefers a non-bold body run).
 */
function extractBaseRunProperties(paragraphXml: string): string {
  const rPrMatches = [...paragraphXml.matchAll(/<w:rPr>[\s\S]*?<\/w:rPr>/g)].map((m) => m[0]);
  if (rPrMatches.length === 0) return '';
  // Prefer a body run (one without <w:b/>) if available, otherwise strip bold from first
  const nonBold = rPrMatches.find((r) => !/<w:b(?:\s*\/>|\s+[^>]*\/>)/.test(r));
  if (nonBold) return nonBold;
  return rPrMatches[0].replace(/<w:b(?:\s*\/>|\s+[^>]*\/>)/g, '');
}

/**
 * Extracts bookmark markers (<w:bookmarkStart .../> and <w:bookmarkEnd .../>)
 * so cross-references and TOC links inside a patched paragraph stay intact.
 */
function extractBookmarks(paragraphXml: string): { starts: string; ends: string } {
  const starts = (paragraphXml.match(/<w:bookmarkStart\b[^>]*\/>/g) || []).join('');
  const ends = (paragraphXml.match(/<w:bookmarkEnd\b[^>]*\/>/g) || []).join('');
  return { starts, ends };
}

/**
 * Extracts live Word cross-reference field sequences (<w:fldChar begin> ... <w:fldChar end>)
 * from a paragraph so that unchanged cross-references (e.g. "Section 8.1") remain live
 * updatable Word REF fields even inside a patched paragraph!
 */
function extractLiveFields(paragraphXml: string): Map<string, string> {
  const fields = new Map<string, string>();
  const fieldRegex =
    /<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>).)*?<w:fldChar\s+w:fldCharType="begin"\/>[\s\S]*?<w:fldChar\s+w:fldCharType="end"\/>(?:(?!<\/w:r>).)*?<\/w:r>/g;
  for (const m of paragraphXml.matchAll(fieldRegex)) {
    const fieldXml = m[0];
    // Extract displayed text between separate and end
    const sepIdx = fieldXml.indexOf('w:fldCharType="separate"');
    if (sepIdx < 0) continue;
    const afterSep = fieldXml.slice(sepIdx);
    const display = [...afterSep.matchAll(/<w:t(?:\s+[^>]*)?>([\s\S]*?)<\/w:t>/g)]
      .map((t) => xmlUnescape(t[1]))
      .join('')
      .trim();
    if (display.length >= 3 && !fields.has(display)) {
      fields.set(display, fieldXml);
    }
  }
  return fields;
}

function emitEqualWithLiveFields(
  text: string,
  baseRPr: string,
  liveFields: Map<string, string>
): string {
  if (liveFields.size === 0) {
    return `<w:r>${baseRPr}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>`;
  }
  for (const [display, fieldXml] of liveFields.entries()) {
    const idx = text.indexOf(display);
    if (idx >= 0) {
      const before = text.slice(0, idx);
      const after = text.slice(idx + display.length);
      const restFields = new Map(liveFields);
      restFields.delete(display);
      return (
        (before ? emitEqualWithLiveFields(before, baseRPr, restFields) : '') +
        fieldXml +
        (after ? emitEqualWithLiveFields(after, baseRPr, restFields) : '')
      );
    }
  }
  return `<w:r>${baseRPr}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>`;
}

function buildTrackedRunsXml(
  tokens: WordDiffToken[],
  author: string,
  isoDate: string,
  revIdRef: { id: number },
  baseRPr: string,
  liveFields: Map<string, string> = new Map()
): string {
  const parts: string[] = [];
  for (const tok of tokens) {
    if (!tok.text) continue;
    const escaped = xmlEscape(tok.text);
    if (tok.type === 'equal') {
      parts.push(emitEqualWithLiveFields(tok.text, baseRPr, liveFields));
    } else if (tok.type === 'del') {
      const id = revIdRef.id++;
      parts.push(
        `<w:del w:id="${id}" w:author="${xmlEscape(author)}" w:date="${isoDate}"><w:r>${baseRPr}<w:delText xml:space="preserve">${escaped}</w:delText></w:r></w:del>`
      );
    } else if (tok.type === 'ins') {
      const id = revIdRef.id++;
      parts.push(
        `<w:ins w:id="${id}" w:author="${xmlEscape(author)}" w:date="${isoDate}"><w:r>${baseRPr}<w:t xml:space="preserve">${escaped}</w:t></w:r></w:ins>`
      );
    }
  }
  return parts.join('');
}

interface PendingComment {
  id: number;
  author: string;
  date: string;
  text: string;
}

function buildCommentElementXml(c: PendingComment): string {
  return `<w:comment w:id="${c.id}" w:author="${xmlEscape(c.author)}" w:date="${c.date}" w:initials="CR"><w:p><w:r><w:t xml:space="preserve">${xmlEscape(
    c.text
  )}</w:t></w:r></w:p></w:comment>`;
}

function ensureContentTypeOverride(ctXml: string, partName: string, contentType: string): string {
  if (ctXml.includes(`PartName="${partName}"`)) return ctXml;
  const override = `  <Override PartName="${partName}" ContentType="${contentType}"/>\n`;
  return ctXml.replace('</Types>', `${override}</Types>`);
}

function ensureDocumentRelationship(
  relsXml: string,
  relType: string,
  target: string
): string {
  if (relsXml.includes(`Type="${relType}"`)) return relsXml;
  let maxRid = 0;
  for (const m of relsXml.matchAll(/Id="rId(\d+)"/g)) {
    const n = Number(m[1]);
    if (n > maxRid) maxRid = n;
  }
  const nextRid = `rId${maxRid + 1}`;
  const rel = `  <Relationship Id="${nextRid}" Type="${relType}" Target="${target}"/>\n`;
  return relsXml.replace('</Relationships>', `${rel}</Relationships>`);
}

async function ensureSettingsAndComments(
  zip: JSZip,
  comments: PendingComment[]
): Promise<void> {
  // 1. Ensure [Content_Types].xml & word/_rels/document.xml.rels exist
  let ctXml = await zip.file('[Content_Types].xml')?.async('string');
  let relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string');
  if (!relsXml) {
    relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;
  }

  // 2. Ensure word/settings.xml has <w:trackRevisions/>
  let settingsXml = await zip.file('word/settings.xml')?.async('string');
  if (!settingsXml) {
    settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:trackRevisions/>
</w:settings>`;
  } else if (!settingsXml.includes('<w:trackRevisions')) {
    settingsXml = settingsXml.replace('</w:settings>', '  <w:trackRevisions/>\n</w:settings>');
  }
  zip.file('word/settings.xml', settingsXml);

  if (ctXml) {
    ctXml = ensureContentTypeOverride(
      ctXml,
      '/word/settings.xml',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
    );
  }
  relsXml = ensureDocumentRelationship(
    relsXml,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
    'settings.xml'
  );

  // 3. Emit or append to word/comments.xml if we have margin comments
  if (comments.length > 0) {
    const commentNodes = comments.map(buildCommentElementXml).join('\n  ');
    let commentsXml = await zip.file('word/comments.xml')?.async('string');
    if (!commentsXml) {
      commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${commentNodes}
</w:comments>`;
    } else {
      commentsXml = commentsXml.replace('</w:comments>', `  ${commentNodes}\n</w:comments>`);
    }
    zip.file('word/comments.xml', commentsXml);

    if (ctXml) {
      ctXml = ensureContentTypeOverride(
        ctXml,
        '/word/comments.xml',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
      );
    }
    relsXml = ensureDocumentRelationship(
      relsXml,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
      'comments.xml'
    );
  }

  if (ctXml) zip.file('[Content_Types].xml', ctXml);
  zip.file('word/_rels/document.xml.rels', relsXml);
}

/**
 * Generates a real Microsoft Word (.docx) file with native OOXML Track Changes
 * (<w:del> / <w:ins>), <w:trackRevisions/>, and real Word Margin Comments (word/comments.xml).
 *
 * - If `sourceDocx` is provided, preserves the original .docx formatting, paragraph
 *   attributes, run fonts (<w:rPr>), bookmarks, and table structure, while patching
 *   only accepted clauses in-place with collision-free revision & comment IDs!
 * - Otherwise, builds a clean, properly structured .docx from scratch.
 */
export async function exportTrackedChangesDocx(opts: {
  title: string;
  counterparty: string;
  playbookName: string;
  clauses: RedlineClause[];
  sourceDocx?: ParsedDocx;
  author?: string;
  /** When true, applies accepted redlines cleanly without Track Changes markup or comments */
  cleanFinal?: boolean;
}): Promise<Buffer> {
  const author = opts.author || 'Contract Redliner';
  const isoDate = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
  const cleanFinal = Boolean(opts.cleanFinal);
  const accepted = opts.clauses.filter(
    (c) => c.decision === 'accepted' && c.originalText !== c.proposedText
  );

  // Path A: In-place patch of an uploaded/local .docx
  if (opts.sourceDocx) {
    const zip = await JSZip.loadAsync(opts.sourceDocx.buffer);
    let docXml = await zip.file('word/document.xml')?.async('string');
    const existingCommentsXml = await zip.file('word/comments.xml')?.async('string');

    if (docXml) {
      const revIdRef = { id: findMaxXmlId(docXml, existingCommentsXml) + 1 };
      const pendingComments: PendingComment[] = [];

      const ranges = findParagraphRanges(docXml);
      const replacements = new Map<number, string>();

      for (const clause of accepted) {
        let targetIdx = clause.paragraphIndex;
        if (targetIdx === undefined || !ranges[targetIdx]) {
          const needle = clause.originalText.slice(0, 60).toLowerCase().replace(/\s+/g, ' ');
          const hit = opts.sourceDocx.paragraphs.find((p) =>
            p.text.toLowerCase().replace(/\s+/g, ' ').includes(needle)
          );
          targetIdx = hit?.index;
        }
        if (targetIdx !== undefined && ranges[targetIdx]) {
          const oldP = ranges[targetIdx].xml;
          const openTagMatch = oldP.match(/^<w:p(?:\s[^>]*)?>/);
          const openTag = openTagMatch ? openTagMatch[0] : '<w:p>';
          const pPrMatch = oldP.match(/<w:pPr[\s\S]*?<\/w:pPr>/);
          const pPr = pPrMatch ? pPrMatch[0] : '';
          const baseRPr = extractBaseRunProperties(oldP);
          const { starts: bmStarts, ends: bmEnds } = extractBookmarks(oldP);
          const liveFields = extractLiveFields(oldP);

          if (cleanFinal) {
            const cleanRun = emitEqualWithLiveFields(clause.proposedText, baseRPr, liveFields);
            replacements.set(targetIdx, `${openTag}${pPr}${bmStarts}${cleanRun}${bmEnds}</w:p>`);
          } else {
            const tokens = computeWordDiff(clause.originalText, clause.proposedText);
            const runsXml = buildTrackedRunsXml(
              tokens,
              author,
              isoDate,
              revIdRef,
              baseRPr,
              liveFields
            );

            let commentStartXml = '';
            let commentEndXml = '';
            if (clause.rationale && clause.rationale.trim()) {
              const commentId = revIdRef.id++;
              pendingComments.push({
                id: commentId,
                author,
                date: isoDate,
                text: `[${clause.playbookRule}] ${clause.rationale}`,
              });
              commentStartXml = `<w:commentRangeStart w:id="${commentId}"/>`;
              commentEndXml = `<w:commentRangeEnd w:id="${commentId}"/><w:r><w:commentReference w:id="${commentId}"/></w:r>`;
            }

            replacements.set(
              targetIdx,
              `${openTag}${pPr}${bmStarts}${commentStartXml}${runsXml}${commentEndXml}${bmEnds}</w:p>`
            );
          }
        }
      }

      // Splice replacements from back to front so byte/char offsets never shift
      let patchedXml = docXml;
      const sortedIndices = Array.from(replacements.keys()).sort((a, b) => b - a);
      for (const idx of sortedIndices) {
        const r = ranges[idx];
        patchedXml =
          patchedXml.slice(0, r.start) + replacements.get(idx)! + patchedXml.slice(r.end);
      }

      zip.file('word/document.xml', patchedXml);
      if (!cleanFinal) {
        await ensureSettingsAndComments(zip, pendingComments);
      }
      return zip.generateAsync({ type: 'nodebuffer' });
    }
  }

  // Path B: Build a clean, valid OOXML .docx from scratch
  const zip = new JSZip();
  const revIdRef = { id: 101 };
  const pendingComments: PendingComment[] = [];

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );

  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`
  );

  const bodyParts: string[] = [];
  bodyParts.push(
    `<w:p><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>${xmlEscape(
      opts.title
    )}${cleanFinal ? ' — Final' : ' — Redlined'}</w:t></w:r></w:p>`
  );
  bodyParts.push(
    `<w:p><w:r><w:rPr><w:color w:val="666666"/><w:sz w:val="20"/></w:rPr><w:t>${xmlEscape(
      `Counterparty: ${opts.counterparty || 'N/A'}  |  Playbook: ${opts.playbookName}  |  Accepted Redlines: ${accepted.length}/${opts.clauses.length}`
    )}</w:t></w:r></w:p>`
  );
  bodyParts.push('<w:p/>');

  for (const clause of opts.clauses) {
    const headingText = `${clause.section ? `${clause.section} — ` : ''}${clause.heading}`;
    bodyParts.push(
      `<w:p><w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>${xmlEscape(
        headingText
      )}</w:t></w:r></w:p>`
    );

    if (clause.decision === 'accepted' && clause.originalText !== clause.proposedText) {
      if (cleanFinal) {
        bodyParts.push(
          `<w:p><w:r><w:t xml:space="preserve">${xmlEscape(clause.proposedText)}</w:t></w:r></w:p>`
        );
      } else {
        const tokens = computeWordDiff(clause.originalText, clause.proposedText);
        const runsXml = buildTrackedRunsXml(tokens, author, isoDate, revIdRef, '');
        const commentId = revIdRef.id++;
        pendingComments.push({
          id: commentId,
          author,
          date: isoDate,
          text: `[${clause.playbookRule}] ${clause.rationale}`,
        });
        bodyParts.push(
          `<w:p><w:commentRangeStart w:id="${commentId}"/>${runsXml}<w:commentRangeEnd w:id="${commentId}"/><w:r><w:commentReference w:id="${commentId}"/></w:r></w:p>`
        );
      }
    } else {
      bodyParts.push(
        `<w:p><w:r><w:t xml:space="preserve">${xmlEscape(clause.originalText)}</w:t></w:r></w:p>`
      );
    }

    bodyParts.push('<w:p/>');
  }

  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyParts.join('\n    ')}
    <w:sectPr/>
  </w:body>
</w:document>`
  );

  if (!cleanFinal) {
    await ensureSettingsAndComments(zip, pendingComments);
  }
  return zip.generateAsync({ type: 'nodebuffer' });
}
