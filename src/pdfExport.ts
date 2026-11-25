import type { RedlineClause } from './types.js';

function pdfEscape(s: string): string {
  // Replace common smart punctuation with ASCII-safe equivalents for Type1 Helvetica
  const ascii = s
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2013\u2014]/g, '-')
    .replace(/\u2026/g, '...')
    .replace(/[^\x20-\x7E]/g, ' ');
  return ascii.replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)');
}

function wrapText(text: string, maxChars: number): string[] {
  const words = text.replace(/\s+/g, ' ').trim().split(' ');
  const lines: string[] = [];
  let cur = '';
  for (const w of words) {
    if (!cur) {
      cur = w;
    } else if (cur.length + 1 + w.length <= maxChars) {
      cur += ' ' + w;
    } else {
      lines.push(cur);
      cur = w;
    }
  }
  if (cur) lines.push(cur);
  return lines;
}

interface PdfLine {
  text: string;
  font: 'F1' | 'F2';
  size: number;
  r: number;
  g: number;
  b: number;
  gapAfter: number;
}

/**
 * Generates a clean, multi-page, zero-dependency PDF Redline Summary Report.
 */
export function exportRedlinePdfReport(opts: {
  title: string;
  counterparty: string;
  playbookName: string;
  clauses: RedlineClause[];
}): Buffer {
  const accepted = opts.clauses.filter((c) => c.decision === 'accepted');
  const lines: PdfLine[] = [];

  const pushWrapped = (
    text: string,
    maxChars: number,
    style: Omit<PdfLine, 'text' | 'gapAfter'>,
    gapAfter = 4
  ) => {
    const wrapped = wrapText(text, maxChars);
    wrapped.forEach((w, idx) => {
      lines.push({
        ...style,
        text: w,
        gapAfter: idx === wrapped.length - 1 ? gapAfter : 2,
      });
    });
  };

  pushWrapped(`${opts.title} - Redline Report`, 64, { font: 'F2', size: 16, r: 0.1, g: 0.12, b: 0.16 }, 6);
  pushWrapped(
    `Counterparty: ${opts.counterparty || 'N/A'}  |  Playbook: ${opts.playbookName}  |  Accepted Edits: ${accepted.length}/${opts.clauses.length}`,
    92,
    { font: 'F1', size: 9.5, r: 0.35, g: 0.38, b: 0.44 },
    14
  );

  for (const c of opts.clauses) {
    const statusLabel =
      c.decision === 'accepted'
        ? 'ACCEPTED REDLINE'
        : c.decision === 'rejected'
          ? 'KEPT ORIGINAL'
          : 'PENDING';

    pushWrapped(
      `${c.section} - ${c.heading}  [${c.riskLevel.toUpperCase()} | ${statusLabel}]`,
      78,
      { font: 'F2', size: 11.5, r: 0.1, g: 0.15, b: 0.25 },
      4
    );

    pushWrapped(
      `Why this matters (${c.playbookRule}): ${c.rationale}`,
      90,
      { font: 'F1', size: 9.5, r: 0.3, g: 0.33, b: 0.38 },
      5
    );

    if (c.decision === 'accepted' && c.originalText !== c.proposedText) {
      pushWrapped(
        `Original: ${c.originalText}`,
        90,
        { font: 'F1', size: 9.5, r: 0.65, g: 0.15, b: 0.15 },
        4
      );
      pushWrapped(
        `Proposed: ${c.proposedText}`,
        90,
        { font: 'F1', size: 9.5, r: 0.08, g: 0.45, b: 0.22 },
        12
      );
    } else {
      pushWrapped(
        `Text (unchanged): ${c.originalText}`,
        90,
        { font: 'F1', size: 9.5, r: 0.22, g: 0.24, b: 0.28 },
        12
      );
    }
  }

  // Paginate onto US Letter (612 x 792 pt, margins 54pt)
  const PAGE_W = 612;
  const PAGE_H = 792;
  const MARGIN_X = 54;
  const TOP_Y = PAGE_H - 56;
  const BOTTOM_Y = 54;

  const pages: string[][] = [];
  let currentOps: string[] = [];
  let y = TOP_Y;

  for (const ln of lines) {
    const lineHeight = ln.size * 1.35;
    if (y - lineHeight < BOTTOM_Y && currentOps.length > 0) {
      pages.push(currentOps);
      currentOps = [];
      y = TOP_Y;
    }
    currentOps.push(
      `BT /${ln.font} ${ln.size} Tf ${ln.r.toFixed(2)} ${ln.g.toFixed(2)} ${ln.b.toFixed(
        2
      )} rg ${MARGIN_X} ${y.toFixed(1)} Td (${pdfEscape(ln.text)}) Tj ET`
    );
    y -= lineHeight + ln.gapAfter;
  }
  if (currentOps.length > 0) pages.push(currentOps);
  if (pages.length === 0) pages.push(['BT /F1 12 Tf 54 736 Td (Empty Report) Tj ET']);

  // Build valid PDF objects:
  // 1: Catalog, 2: Pages, 3: F1 Helvetica, 4: F2 Helvetica-Bold, then pairs of (Page, Content)
  const objs: string[] = [];
  const pageObjNums: number[] = [];
  const contentStart = 5;

  for (let i = 0; i < pages.length; i++) {
    pageObjNums.push(contentStart + i * 2);
  }

  objs.push(`1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n`);
  objs.push(
    `2 0 obj << /Type /Pages /Kids [${pageObjNums
      .map((n) => `${n} 0 R`)
      .join(' ')}] /Count ${pages.length} >> endobj\n`
  );
  objs.push(`3 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj\n`);
  objs.push(`4 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >> endobj\n`);

  pages.forEach((ops, idx) => {
    const pageNum = contentStart + idx * 2;
    const streamNum = pageNum + 1;
    const stream = ops.join('\n');
    objs.push(
      `${pageNum} 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 ${PAGE_W} ${PAGE_H}] /Contents ${streamNum} 0 R /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> >> endobj\n`
    );
    objs.push(
      `${streamNum} 0 obj << /Length ${Buffer.byteLength(stream, 'utf-8')} >> stream\n${stream}\nendstream endobj\n`
    );
  });

  let pdf = '%PDF-1.4\n';
  const offsets: number[] = [];
  for (const o of objs) {
    offsets.push(Buffer.byteLength(pdf, 'utf-8'));
    pdf += o;
  }
  const xrefPos = Buffer.byteLength(pdf, 'utf-8');
  pdf += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n`;
  for (const off of offsets) {
    pdf += String(off).padStart(10, '0') + ' 00000 n \n';
  }
  pdf += `trailer << /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xrefPos}\n%%EOF\n`;

  return Buffer.from(pdf, 'utf-8');
}

export function exportMarkdownMemo(opts: {
  title: string;
  counterparty: string;
  playbookName: string;
  clauses: RedlineClause[];
}): Buffer {
  const accepted = opts.clauses.filter((c) => c.decision === 'accepted');
  const md: string[] = [
    `# ${opts.title} — Redline Memo`,
    '',
    `- **Counterparty:** ${opts.counterparty || 'N/A'}`,
    `- **Playbook:** ${opts.playbookName}`,
    `- **Accepted Redlines:** ${accepted.length} of ${opts.clauses.length} reviewed clauses`,
    '',
    '---',
    '',
  ];

  for (const c of opts.clauses) {
    const badge =
      c.decision === 'accepted'
        ? '✅ Accepted'
        : c.decision === 'rejected'
          ? '⏭ Kept Original'
          : '⏳ Pending';
    md.push(`## ${c.section} — ${c.heading} (${c.riskLevel.toUpperCase()} · ${badge})`);
    md.push('');
    md.push(`> **Why this matters (${c.playbookRule}):** ${c.rationale}`);
    md.push('');
    if (c.decision === 'accepted' && c.originalText !== c.proposedText) {
      md.push(`**Original:** ${c.originalText}`);
      md.push('');
      md.push(`**Proposed Redline:** ${c.proposedText}`);
    } else {
      md.push(`**Text (unchanged):** ${c.originalText}`);
    }
    md.push('');
  }

  return Buffer.from(md.join('\n'), 'utf-8');
}
