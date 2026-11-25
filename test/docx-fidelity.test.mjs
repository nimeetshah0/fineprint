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

const TMP_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'docx-fidelity-'));

async function validateDocxZip(buf, label) {
  const zip = await JSZip.loadAsync(buf);
  const xmlFiles = Object.keys(zip.files).filter(
    (name) => !zip.files[name].dir && (name.endsWith('.xml') || name.endsWith('.rels'))
  );
  assert.ok(xmlFiles.includes('word/document.xml'), `${label}: missing word/document.xml`);

  // Validate EVERY XML file in the archive with xmllint
  for (const name of xmlFiles) {
    const content = await zip.file(name).async('string');
    const tmpXml = path.join(TMP_DIR, name.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpXml, content, 'utf-8');
    try {
      execFileSync('xmllint', ['--noout', tmpXml], { stdio: 'pipe' });
    } catch (err) {
      throw new Error(
        `${label}: XML well-formedness failed for ${name}: ${err.stderr?.toString() || err.message}`
      );
    }
  }

  // Also test Apple's native Office importer via textutil
  const tmpDocx = path.join(TMP_DIR, `${label.replace(/[^a-z0-9]/gi, '_')}.docx`);
  fs.writeFileSync(tmpDocx, buf);
  const txtOut = execFileSync('textutil', ['-convert', 'txt', '-stdout', tmpDocx], {
    encoding: 'utf-8',
  });
  assert.ok(txtOut.trim().length > 0, `${label}: textutil produced empty output`);
  return { zip, txtOut };
}

/**
 * Builds a realistic law-firm-style .docx containing:
 * - Styled Georgia 11pt runs + bold defined terms
 * - Numbered paragraph properties (<w:pPr><w:numPr>)
 * - Bookmarks (<w:bookmarkStart> / <w:bookmarkEnd>) for cross-references
 * - A Pricing & SLA Table (<w:tbl>) with cells
 * - Existing counterparty Track Changes (<w:del w:id="100"> / <w:ins w:id="101">)
 * - Special characters (&, quotes, em-dash)
 * - NO word/settings.xml initially (common in Google Docs / Pandoc exports)
 */
async function buildTortureContractDocx() {
  const zip = new JSZip();

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

  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000001">
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:sz w:val="28"/></w:rPr><w:t>MASTER SERVICES AGREEMENT — ACME &amp; CO.</w:t></w:r>
    </w:p>
    <w:p/>
    <w:p w14:paraId="0000000A">
      <w:r><w:t>Preamble with embedded box</w:t></w:r>
      <w:txbxContent><w:p><w:r><w:t>Nested textbox paragraph</w:t></w:r></w:p></w:txbxContent>
    </w:p>

    <w:p w14:paraId="00000002">
      <w:pPr><w:spacing w:after="160"/></w:pPr>
      <w:bookmarkStart w:id="50" w:name="sec_liability"/>
      <w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">8.1 Limitation of Liability. </w:t></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">Customer shall remain liable without limitation, while Vendor aggregate liability shall not exceed $100.</w:t></w:r>
      <w:bookmarkEnd w:id="50"/>
    </w:p>

    <w:p w14:paraId="00000003">
      <w:pPr><w:spacing w:after="160"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">3.3 Auto-Renewal. This Agreement renews for </w:t></w:r>
      <w:del w:id="100" w:author="Counterparty Counsel" w:date="2026-11-20T10:00:00Z"><w:r><w:delText xml:space="preserve">one-year</w:delText></w:r></w:del>
      <w:ins w:id="101" w:author="Counterparty Counsel" w:date="2026-11-20T10:00:00Z"><w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">three-year</w:t></w:r></w:ins>
      <w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve"> terms unless 90 days notice is given.</w:t></w:r>
    </w:p>

    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Order Term</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="20"/></w:rPr><w:t>Net 7 days from invoice date</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>

    <w:sectPr/>
  </w:body>
</w:document>`
  );

  const filePath = path.join(TMP_DIR, 'torture_msa.docx');
  const buf = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(filePath, buf);
  return filePath;
}

const results = [];
function check(name, fn) {
  try {
    fn();
    results.push({ name, pass: true });
  } catch (err) {
    results.push({ name, pass: false, err: err.message });
  }
}

const torturePath = await buildTortureContractDocx();
const parsed = await parseDocxFile(torturePath);

// Build redlines targeting:
// 1. Liability clause (styled Georgia + bookmark)
// 2. Auto-Renewal clause (already has counterparty w:id="100" & w:id="101"!)
// 3. Pricing table cell ("Net 7 days from invoice date" -> "Net 30 days from receipt of undisputed invoice")
const liabP = parsed.paragraphs.find((p) => p.text.includes('Limitation of Liability'));
const renewP = parsed.paragraphs.find((p) => p.text.includes('Auto-Renewal'));
const tableP = parsed.paragraphs.find((p) => p.text.includes('Net 7 days'));

const clauses = [
  {
    id: 'c-1',
    section: 'Section 8.1',
    heading: 'Limitation of Liability',
    originalText: liabP.text,
    proposedText:
      '8.1 Limitation of Liability. Each party aggregate liability shall not exceed 12 months of fees paid.',
    riskLevel: 'dealbreaker',
    playbookRule: 'Limitation of Liability',
    rationale: 'Standard mutual 12-month fee cap per company playbook.',
    decision: 'accepted',
    paragraphIndex: liabP.index,
    diffTokens: computeWordDiff(
      liabP.text,
      '8.1 Limitation of Liability. Each party aggregate liability shall not exceed 12 months of fees paid.'
    ),
  },
  {
    id: 'c-2',
    section: 'Section 3.3',
    heading: 'Auto-Renewal',
    originalText: renewP.text,
    proposedText:
      '3.3 Auto-Renewal. This Agreement renews for one-year terms unless 30 days notice is given.',
    riskLevel: 'negotiate',
    playbookRule: 'Auto-Renewal',
    rationale: 'Shorten renewal term back to 1 year with 30-day notice.',
    decision: 'accepted',
    paragraphIndex: renewP.index,
    diffTokens: computeWordDiff(
      renewP.text,
      '3.3 Auto-Renewal. This Agreement renews for one-year terms unless 30 days notice is given.'
    ),
  },
  {
    id: 'c-3',
    section: 'Order Form Table',
    heading: 'Payment Terms',
    originalText: tableP.text,
    proposedText: 'Net 30 days from receipt of undisputed invoice',
    riskLevel: 'negotiate',
    playbookRule: 'Payment Terms',
    rationale: 'Move from Net 7 to Net 30.',
    decision: 'accepted',
    paragraphIndex: tableP.index,
    diffTokens: computeWordDiff(
      tableP.text,
      'Net 30 days from receipt of undisputed invoice'
    ),
  },
];

const exportedBuf = await exportTrackedChangesDocx({
  title: 'Torture MSA',
  counterparty: 'Acme & Co.',
  playbookName: 'Startup Vendor MSA',
  clauses,
  sourceDocx: parsed,
});

const { zip } = await validateDocxZip(exportedBuf, 'TortureMSA');
const docXml = await zip.file('word/document.xml').async('string');
const settingsXml = await zip.file('word/settings.xml')?.async('string');
const commentsXml = await zip.file('word/comments.xml')?.async('string');
const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string');
const ctXml = await zip.file('[Content_Types].xml')?.async('string');

check('1. Parse reads current view (excludes counterparty <w:delText>)', () => {
  assert.ok(renewP.text.includes('three-year'), `Expected three-year in: ${renewP.text}`);
  assert.ok(!renewP.text.includes('one-year'), `Should not include deleted text in: ${renewP.text}`);
});

check('2. Preserves run font styling (<w:rFonts w:ascii="Georgia">) on patched clauses', () => {
  const p2 = docXml.match(/<w:p\b[^>]*>[\s\S]*?Limitation of Liability[\s\S]*?<\/w:p>/)?.[0] || '';
  assert.ok(
    p2.includes('w:ascii="Georgia"'),
    'Georgia font <w:rPr> was wiped out when patching Section 8.1!'
  );
});

check('3. Preserves bookmarks (<w:bookmarkStart w:name="sec_liability"/>) for cross-references', () => {
  assert.ok(
    docXml.includes('w:name="sec_liability"'),
    'Bookmark sec_liability was destroyed when patching Section 8.1!'
  );
});

check('4. Preserves table structure (<w:tbl>) while redlining cell', () => {
  assert.ok(docXml.includes('<w:tbl>') && docXml.includes('</w:tbl>'), 'Table broken');
  const tableMatch = docXml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)?.[0] || '';
  assert.ok(tableMatch.includes('<w:ins '), 'Table cell missing <w:ins> track change');
  assert.ok(tableMatch.includes('>30<'), 'Table cell missing inserted "30"');
  assert.ok(tableMatch.includes('<w:delText xml:space="preserve">7</w:delText>'), 'Table cell missing deleted "7"');
});

check('5. No duplicate w:id collisions with existing counterparty revisions', () => {
  const ids = [...docXml.matchAll(/<w:(?:del|ins)\b[^>]*w:id="(\d+)"/g)].map((m) => m[1]);
  const unique = new Set(ids);
  assert.equal(ids.length, unique.size, `Duplicate revision w:id detected: ${ids.join(', ')}`);
  assert.ok(!ids.includes('100') && !ids.includes('101'), 'Reused counterparty revision IDs 100/101!');
});

check('6. Injects <w:trackRevisions/> into word/settings.xml + registers rels & content types', () => {
  assert.ok(settingsXml && settingsXml.includes('<w:trackRevisions/>'), 'Missing word/settings.xml <w:trackRevisions/>');
  assert.ok(relsXml.includes('relationships/settings'), 'Missing settings relationship in document.xml.rels');
  assert.ok(ctXml.includes('/word/settings.xml'), 'Missing settings override in [Content_Types].xml');
});

check('7. Emits real Microsoft Word margin comments (word/comments.xml + range markers)', () => {
  assert.ok(commentsXml && commentsXml.includes('<w:comment '), 'Missing word/comments.xml');
  assert.ok(docXml.includes('<w:commentRangeStart '), 'Missing <w:commentRangeStart> in document.xml');
  assert.ok(docXml.includes('<w:commentReference '), 'Missing <w:commentReference> in document.xml');
  assert.ok(relsXml.includes('relationships/comments'), 'Missing comments relationship in document.xml.rels');
  assert.ok(ctXml.includes('/word/comments.xml'), 'Missing comments override in [Content_Types].xml');
});

check('8. Survives self-closing <w:p/> spacers and nested <w:txbxContent><w:p> without shifting or corruption', () => {
  assert.ok(docXml.includes('<w:p/>'), 'Self-closing <w:p/> spacer lost');
  assert.ok(docXml.includes('Nested textbox paragraph'), 'Nested textbox paragraph corrupted');
});

let failed = 0;
for (const r of results) {
  console.log(`${r.pass ? 'PASS' : 'FAIL'} | ${r.name}${r.err ? `\n       -> ${r.err}` : ''}`);
  if (!r.pass) failed++;
}
console.log(`\nSummary: ${results.length - failed}/${results.length} passed`);
process.exit(failed > 0 ? 1 : 0);
