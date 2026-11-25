import type { Playbook } from './types.js';

export const SKILL_URI = 'skill://fineprint/SKILL.md';
export const SKILL_PLAYBOOKS_URI = 'skill://fineprint/references/playbooks.md';

export function renderFineprintSkillMd(): string {
  return `---
name: fineprint-contract-review
description: >-
  Review contracts (.docx, .pdf, Google Docs, or pasted text) against startup,
  agency, or NDA playbooks, render interactive Before/After redline cards in
  Fineprint, and export native Microsoft Word (.docx) Track Changes with margin
  comments. Drafting aid only — not legal advice.
---

# Fineprint Contract Review Workflow

Use this skill whenever the user asks to review, redline, or check a contract
(MSA, SOW, NDA, vendor agreement, or Google Doc) using Fineprint.

## Step 1: Load the contract
- **Google Docs URL** (\`https://docs.google.com/document/d/...\`) → call \`read_google_doc\` with the URL.
- **Local \`.docx\` or \`.pdf\` path** (when running locally) → call \`read_contract\` with the path.
- **Uploaded/pasted text** → read the clauses directly from the conversation.

## Step 2: Load the matching playbook
Call \`list_playbooks\` (or read \`${SKILL_PLAYBOOKS_URI}\`) and select the closest playbook:
- \`startup-vendor-msa\` — SaaS vendor or enterprise supplier MSAs
- \`agency-client-sow\` — Agency / consultant client SOWs & MSAs
- \`mutual-nda\` — Mutual NDAs

## Step 3: Analyze and open Fineprint
For each material clause (dealbreakers, negotiable terms, and key standard clauses):
1. Quote the **exact** \`originalText\` verbatim (preserving wording so word-level Track Changes diff cleanly).
2. Draft a balanced, commercially reasonable \`proposedText\` aligned with the playbook's position (or fallback if appropriate).
3. Assign \`riskLevel\` (\`dealbreaker\`, \`negotiate\`, or \`standard\`), \`playbookRule\`, and a crisp 1–2 sentence plain-English \`rationale\`.
4. Include \`paragraphIndex\` (from \`read_google_doc\` or \`read_contract\`) and \`sessionId\` / \`googleDocUrl\` so exported \`.docx\` Track Changes preserve the original document's formatting, tables, bookmarks, and live cross-references.
5. Call \`open_redliner\` to render the interactive Material 3 workbench.

## Step 4: Softening & Export
- When the user asks to soften a clause, apply the playbook's fallback position and call \`open_redliner\` with the updated wording.
- Remind the user that Fineprint is a drafting and formatting aid only (not legal advice), and that exported Track Changes \`.docx\` files re-import cleanly into both Microsoft Word and Google Docs (as native Suggestions & Comments).
`;
}

export function renderPlaybooksReferenceMd(playbooks: Playbook[]): string {
  const sections = playbooks.map((pb) => {
    const rules = pb.rules
      .map(
        (r) =>
          `- **${r.topic}** (\`${r.severity}\`): ${r.position}\n  - *Fallback:* ${r.fallback}`
      )
      .join('\n');
    return `## ${pb.name} (\`${pb.id}\`)\n*Audience:* ${pb.audience}\n\n${pb.description}\n\n${rules}`;
  });
  return `# Fineprint Playbooks Reference\n\n${sections.join('\n\n---\n\n')}\n`;
}

async function sha256Hex(text: string): Promise<string> {
  const bytes = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(digest), (b) => b.toString(16).padStart(2, '0')).join('');
}

export async function buildSkillManifest(playbooks: Playbook[]) {
  const skillMd = renderFineprintSkillMd();
  const playbooksMd = renderPlaybooksReferenceMd(playbooks);
  const [skillHash, pbHash] = await Promise.all([
    sha256Hex(skillMd),
    sha256Hex(playbooksMd),
  ]);

  return {
    uri: SKILL_URI,
    frontmatter: {
      name: 'fineprint-contract-review',
      description:
        'Review contracts (.docx, .pdf, Google Docs, or pasted text) against Fineprint playbooks and export native Word Track Changes (.docx) with margin comments.',
    },
    resources: [
      {
        uri: SKILL_URI,
        mimeType: 'text/markdown',
        digest: `sha256:${skillHash}`,
        size: new TextEncoder().encode(skillMd).byteLength,
      },
      {
        uri: SKILL_PLAYBOOKS_URI,
        mimeType: 'text/markdown',
        digest: `sha256:${pbHash}`,
        size: new TextEncoder().encode(playbooksMd).byteLength,
      },
    ],
  };
}
