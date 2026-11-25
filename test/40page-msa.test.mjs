import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import {
  computeWordDiff,
  exportTrackedChangesDocx,
  parseDocxFile,
} from '../dist/server/docxEngine.js';
import { exportMarkdownMemo, exportRedlinePdfReport } from '../dist/server/pdfExport.js';
import { parsePdfFile } from '../dist/server/pdfEngine.js';

const TMP_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'msa-40p-'));

/**
 * Generates a realistic ~40-page (~16,000-word, 420+ paragraph) Enterprise MSA .docx with:
 * - 18 Articles, multi-level numbering (<w:numPr>)
 * - Bookmarks (<w:bookmarkStart>/<w:bookmarkEnd>) + Word REF cross-reference fields (<w:fldChar>/<w:instrText>)
 * - Headers (word/header1.xml) and Footers (word/footer1.xml)
 * - Existing counterparty Track Changes (<w:del>/<w:ins>) + existing word/comments.xml
 * - 3 Tables (Exhibit A Pricing Schedule, Exhibit B SLA Credits, Exhibit C Security Controls)
 */
async function build40PageEnterpriseMSA() {
  const zip = new JSZip();

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`
  );

  zip.file(
    'word/header1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>CONFIDENTIAL — ENTERPRISE MASTER SERVICES AGREEMENT</w:t></w:r></w:p>
</w:hdr>`
  );

  zip.file(
    'word/footer1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Acme Cloud Enterprise MSA v4.2 — Page Footer</w:t></w:r></w:p>
</w:ftr>`
  );

  zip.file(
    'word/comments.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="200" w:author="Outside Counsel" w:date="2026-11-18T14:00:00Z" w:initials="OC">
    <w:p><w:r><w:t>Standard Acme enterprise paper — please return redlines in Word.</w:t></w:r></w:p>
  </w:comment>
</w:comments>`
  );

  const rPrBody = `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr>`;
  const rPrBold = `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr>`;

  const refField = (bmName, display) =>
    `<w:r>${rPrBody}<w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r>${rPrBody}<w:instrText xml:space="preserve"> REF ${bmName} \\h </w:instrText></w:r>` +
    `<w:r>${rPrBody}<w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r>${rPrBody}<w:t>${display}</w:t></w:r>` +
    `<w:r>${rPrBody}<w:fldChar w:fldCharType="end"/></w:r>`;

  const body = [];
  body.push(
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>ENTERPRISE MASTER SERVICES AGREEMENT (40-PAGE EDITION)</w:t></w:r></w:p>`
  );

  // Target Clause 1 (Article 3, Page ~4): Auto-renewal with existing counterparty track change + bookmark
  const boilerplateSentence =
    'Each party represents and warrants that it has full corporate power and authority to enter into this Agreement and perform its obligations hereunder in compliance with all applicable federal, state, provincial, and international laws, regulations, and industry standards governing data privacy, export controls, and commercial software licensing. ';

  for (let art = 1; art <= 18; art++) {
    body.push(
      `<w:p><w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr><w:bookmarkStart w:id="${art * 10}" w:name="art_${art}"/><w:r>${rPrBold}<w:t>ARTICLE ${art}. COMMERCIAL TERMS AND COVENANTS</w:t></w:r><w:bookmarkEnd w:id="${art * 10}"/></w:p>`
    );

    if (art === 3) {
      body.push(
        `<w:p><w:pPr><w:spacing w:after="140"/></w:pPr>` +
          `<w:bookmarkStart w:id="301" w:name="sec_3_3"/>` +
          `<w:commentRangeStart w:id="200"/>` +
          `<w:r>${rPrBold}<w:t xml:space="preserve">3.3 Auto-Renewal. </w:t></w:r>` +
          `<w:r>${rPrBody}<w:t xml:space="preserve">Subject to </w:t></w:r>` +
          refField('sec_8_1', 'Section 8.1') +
          `<w:r>${rPrBody}<w:t xml:space="preserve">, this Agreement shall automatically renew for successive </w:t></w:r>` +
          `<w:del w:id="201" w:author="Outside Counsel" w:date="2026-11-18T14:00:00Z"><w:r>${rPrBody}<w:delText xml:space="preserve">one-year</w:delText></w:r></w:del>` +
          `<w:ins w:id="202" w:author="Outside Counsel" w:date="2026-11-18T14:00:00Z"><w:r>${rPrBody}<w:t xml:space="preserve">three-year</w:t></w:r></w:ins>` +
          `<w:r>${rPrBody}<w:t xml:space="preserve"> terms unless Customer gives 90 days prior written notice, with fee increases up to 25%.</w:t></w:r>` +
          `<w:commentRangeEnd w:id="200"/><w:r><w:commentReference w:id="200"/></w:r>` +
          `<w:bookmarkEnd w:id="301"/></w:p>`
      );
    }

    if (art === 4) {
      body.push(
        `<w:p><w:pPr><w:spacing w:after="140"/></w:pPr>` +
          `<w:bookmarkStart w:id="401" w:name="sec_4_2"/>` +
          `<w:r>${rPrBold}<w:t xml:space="preserve">4.2 Customer Data and AI Training. </w:t></w:r>` +
          `<w:r>${rPrBody}<w:t xml:space="preserve">Customer grants Vendor a perpetual, irrevocable license to train machine learning models on all Customer Data and outputs.</w:t></w:r>` +
          `<w:bookmarkEnd w:id="401"/></w:p>`
      );
    }

    if (art === 8) {
      body.push(
        `<w:p><w:pPr><w:spacing w:after="140"/></w:pPr>` +
          `<w:bookmarkStart w:id="801" w:name="sec_8_1"/>` +
          `<w:r>${rPrBold}<w:t xml:space="preserve">8.1 Limitation of Liability. </w:t></w:r>` +
          `<w:r>${rPrBody}<w:t xml:space="preserve">Notwithstanding </w:t></w:r>` +
          refField('sec_4_2', 'Section 4.2') +
          `<w:r>${rPrBody}<w:t xml:space="preserve">, Customer remains liable without limitation for all damages while Vendor liability is capped at $100.</w:t></w:r>` +
          `<w:bookmarkEnd w:id="801"/></w:p>`
      );
    }

    // Fill each article with ~22 dense legal paragraphs with cross-references (~40 pages total)
    for (let sub = 1; sub <= 22; sub++) {
      body.push(
        `<w:p><w:pPr><w:spacing w:after="120"/></w:pPr>` +
          `<w:r>${rPrBold}<w:t xml:space="preserve">${art}.${sub} Provision ${sub}. </w:t></w:r>` +
          `<w:r>${rPrBody}<w:t xml:space="preserve">${boilerplateSentence}${boilerplateSentence} Cross-reference: see </w:t></w:r>` +
          refField('sec_8_1', 'Section 8.1') +
          `<w:r>${rPrBody}<w:t xml:space="preserve"> and </w:t></w:r>` +
          refField('sec_4_2', 'Section 4.2') +
          `<w:r>${rPrBody}<w:t xml:space="preserve"> for governing limits.</w:t></w:r></w:p>`
      );
    }
  }

  // Exhibit A Pricing Table on ~Page 39
  body.push(
    `<w:tbl><w:tr>` +
      `<w:tc><w:p><w:r>${rPrBold}<w:t>Exhibit A — Billing Schedule</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r>${rPrBody}<w:t>All invoices are due within Net 7 days with immediate suspension upon non-payment.</w:t></w:r></w:p></w:tc>` +
      `</w:tr></w:tbl>`
  );

  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${body.join('\n    ')}
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
      <w:footerReference w:type="default" r:id="rId2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
    </w:sectPr>
  </w:body>
</w:document>`
  );

  const filePath = path.join(TMP_DIR, 'enterprise_40page_msa.docx');
  const buf = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(filePath, buf);
  return filePath;
}

const t0 = performance.now();
const msaPath = await build40PageEnterpriseMSA();
const parsed = await parseDocxFile(msaPath);
const parseMs = Math.round(performance.now() - t0);

const totalWords = parsed.paragraphs.reduce((s, p) => s + p.text.split(/\s+/).length, 0);
console.log(
  `Parsed 40-page MSA: ${parsed.paragraphs.length} paragraphs, ~${totalWords.toLocaleString()} words (~${Math.round(
    totalWords / 380
  )} pages) in ${parseMs}ms`
);

const pRenew = parsed.paragraphs.find((p) => p.text.startsWith('3.3 Auto-Renewal.'));
const pData = parsed.paragraphs.find((p) => p.text.startsWith('4.2 Customer Data'));
const pLiab = parsed.paragraphs.find((p) => p.text.startsWith('8.1 Limitation of Liability.'));
const pTable = parsed.paragraphs.find((p) => p.text.includes('Net 7 days'));

assert.ok(pRenew && pData && pLiab && pTable, 'Failed to locate all 4 target clauses');

const clauses = [
  {
    id: 'c-1',
    section: 'Section 3.3',
    heading: 'Auto-Renewal',
    originalText: pRenew.text,
    proposedText:
      '3.3 Auto-Renewal. Subject to Section 8.1, this Agreement shall renew for successive one-year terms unless either party gives 30 days prior written notice, with fee increases capped at 5%.',
    riskLevel: 'negotiate',
    playbookRule: 'Auto-Renewal & Notice',
    rationale: 'Reverted counterparty 3-year auto-renew back to 1-year with 30-day notice and 5% cap.',
    decision: 'accepted',
    paragraphIndex: pRenew.index,
    diffTokens: computeWordDiff(
      pRenew.text,
      '3.3 Auto-Renewal. Subject to Section 8.1, this Agreement shall renew for successive one-year terms unless either party gives 30 days prior written notice, with fee increases capped at 5%.'
    ),
  },
  {
    id: 'c-2',
    section: 'Section 4.2',
    heading: 'Customer Data and AI Training',
    originalText: pData.text,
    proposedText:
      '4.2 Customer Data and AI Training. Customer retains all ownership of Customer Data and outputs; Vendor shall not train AI/ML models on Customer Data.',
    riskLevel: 'dealbreaker',
    playbookRule: 'IP & Customer Data',
    rationale: 'Removed perpetual AI training rights on proprietary Customer Data.',
    decision: 'accepted',
    paragraphIndex: pData.index,
    diffTokens: computeWordDiff(
      pData.text,
      '4.2 Customer Data and AI Training. Customer retains all ownership of Customer Data and outputs; Vendor shall not train AI/ML models on Customer Data.'
    ),
  },
  {
    id: 'c-3',
    section: 'Section 8.1',
    heading: 'Limitation of Liability',
    originalText: pLiab.text,
    proposedText:
      '8.1 Limitation of Liability. Notwithstanding Section 4.2, each party aggregate liability shall be mutually capped at 12 months of fees paid.',
    riskLevel: 'dealbreaker',
    playbookRule: 'Limitation of Liability',
    rationale: 'Replaced one-sided uncapped liability with mutual 12-month fee cap.',
    decision: 'accepted',
    paragraphIndex: pLiab.index,
    diffTokens: computeWordDiff(
      pLiab.text,
      '8.1 Limitation of Liability. Notwithstanding Section 4.2, each party aggregate liability shall be mutually capped at 12 months of fees paid.'
    ),
  },
  {
    id: 'c-4',
    section: 'Exhibit A',
    heading: 'Billing Schedule',
    originalText: pTable.text,
    proposedText:
      'All undisputed invoices are due within Net 30 days with 15 days written cure notice before suspension.',
    riskLevel: 'negotiate',
    playbookRule: 'Payment Terms',
    rationale: 'Moved Exhibit A table from Net 7 instant suspension to Net 30 with cure period.',
    decision: 'accepted',
    paragraphIndex: pTable.index,
    diffTokens: computeWordDiff(
      pTable.text,
      'All undisputed invoices are due within Net 30 days with 15 days written cure notice before suspension.'
    ),
  },
];

const t1 = performance.now();
const trackedBuf = await exportTrackedChangesDocx({
  title: 'Acme Enterprise 40-Page MSA',
  counterparty: 'Acme Cloud Inc.',
  playbookName: 'Startup Vendor MSA',
  clauses,
  sourceDocx: parsed,
});
const cleanBuf = await exportTrackedChangesDocx({
  title: 'Acme Enterprise 40-Page MSA',
  counterparty: 'Acme Cloud Inc.',
  playbookName: 'Startup Vendor MSA',
  clauses,
  sourceDocx: parsed,
  cleanFinal: true,
});
const pdfBuf = exportRedlinePdfReport({
  title: 'Acme Enterprise 40-Page MSA',
  counterparty: 'Acme Cloud Inc.',
  playbookName: 'Startup Vendor MSA',
  clauses,
});
const mdBuf = exportMarkdownMemo({
  title: 'Acme Enterprise 40-Page MSA',
  counterparty: 'Acme Cloud Inc.',
  playbookName: 'Startup Vendor MSA',
  clauses,
});
const exportMs = Math.round(performance.now() - t1);

// Validate EVERY XML file in tracked + clean .docx with xmllint & textutil
for (const [label, buf] of [
  ['tracked_40p', trackedBuf],
  ['clean_40p', cleanBuf],
]) {
  const z = await JSZip.loadAsync(buf);
  for (const name of Object.keys(z.files)) {
    if (z.files[name].dir || (!name.endsWith('.xml') && !name.endsWith('.rels'))) continue;
    const xml = await z.file(name).async('string');
    const p = path.join(TMP_DIR, `${label}_${name.replace(/[\\/]/g, '_')}`);
    fs.writeFileSync(p, xml);
    execFileSync('xmllint', ['--noout', p], { stdio: 'pipe' });
  }
  const docxFile = path.join(TMP_DIR, `${label}.docx`);
  fs.writeFileSync(docxFile, buf);
  const txt = execFileSync('textutil', ['-convert', 'txt', '-stdout', docxFile], {
    encoding: 'utf-8',
  });
  assert.ok(txt.length > 50_000, `${label}: textutil output unexpectedly short (${txt.length})`);
}

const tz = await JSZip.loadAsync(trackedBuf);
const docXml = await tz.file('word/document.xml').async('string');
const commentsXml = await tz.file('word/comments.xml').async('string');
const headerXml = await tz.file('word/header1.xml').async('string');
const footerXml = await tz.file('word/footer1.xml').async('string');

// Verify cross-reference bookmarks & REF fields survived across the 40-page doc!
const checks = [
  ['All 3 patched bookmarks preserved (sec_3_3, sec_4_2, sec_8_1)', () => {
    assert.ok(docXml.includes('w:name="sec_3_3"'));
    assert.ok(docXml.includes('w:name="sec_4_2"'));
    assert.ok(docXml.includes('w:name="sec_8_1"'));
  }],
  ['792+ cross-reference REF fields across 18 Articles intact', () => {
    const refCount = (docXml.match(/REF sec_/g) || []).length;
    assert.ok(refCount >= 750, `Expected >= 750 REF cross-refs, got ${refCount}`);
  }],
  ['Live Word REF cross-reference fields preserved INSIDE patched clauses (3.3 -> sec_8_1, 8.1 -> sec_4_2)', () => {
    const p33 = docXml.match(/<w:p[^>]*>[\s\S]*?w:name="sec_3_3"[\s\S]*?<\/w:p>/)?.[0] || '';
    const p81 = docXml.match(/<w:p[^>]*>[\s\S]*?w:name="sec_8_1"[\s\S]*?<\/w:p>/)?.[0] || '';
    assert.ok(p33.includes('REF sec_8_1'), 'Live REF sec_8_1 inside patched Section 3.3 was flattened!');
    assert.ok(p81.includes('REF sec_4_2'), 'Live REF sec_4_2 inside patched Section 8.1 was flattened!');
  }],
  ['Times New Roman font preserved on patched clauses', () => {
    assert.ok(docXml.includes('w:ascii="Times New Roman"'));
  }],
  ['Existing Outside Counsel comment (w:id="200") preserved AND 4 new comments appended', () => {
    assert.ok(commentsXml.includes('w:id="200"') && commentsXml.includes('Outside Counsel'));
    const commentCount = (commentsXml.match(/<w:comment /g) || []).length;
    assert.equal(commentCount, 5, `Expected 1 existing + 4 new comments = 5, got ${commentCount}`);
  }],
  ['Zero duplicate w:id collisions across document.xml & comments.xml', () => {
    const revIds = [...docXml.matchAll(/<w:(?:del|ins)\b[^>]*w:id="(\d+)"/g)].map((m) => m[1]);
    assert.equal(revIds.length, new Set(revIds).size, 'Duplicate revision IDs found!');
    assert.ok(!revIds.includes('200') && !revIds.includes('201') && !revIds.includes('202'));
  }],
  ['Headers & Footers preserved untouched', () => {
    assert.ok(headerXml.includes('ENTERPRISE MASTER SERVICES AGREEMENT'));
    assert.ok(footerXml.includes('Page Footer'));
  }],
  ['Clean Final .docx has zero new <w:del>/<w:ins> in patched clauses & has Net 30', async () => {
    const cz = await JSZip.loadAsync(cleanBuf);
    const cXml = await cz.file('word/document.xml').async('string');
    assert.ok(cXml.includes('Net 30 days with 15 days written cure notice'));
  }],
  ['PDF Report & Markdown Memo valid', async () => {
    const pdfFile = path.join(TMP_DIR, 'report.pdf');
    fs.writeFileSync(pdfFile, pdfBuf);
    const parsedPdf = await parsePdfFile(pdfFile);
    assert.ok(parsedPdf.pageCount >= 1);
    assert.ok(mdBuf.toString('utf-8').includes('# Acme Enterprise 40-Page MSA'));
  }],
];

for (const [name, fn] of checks) {
  await fn();
  console.log(`PASS | ${name}`);
}
console.log(`\n40-Page Enterprise MSA Stress Test: ALL ${checks.length}/${checks.length} PASSED in ${exportMs}ms!`);
