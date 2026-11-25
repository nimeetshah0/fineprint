import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { Playbook, RedlineClause } from './types.js';
import { computeWordDiff } from './docxEngine.js';
import builtInPlaybooksJson from './playbooks.json' with { type: 'json' };

export const BUILT_IN_PLAYBOOKS: Playbook[] = builtInPlaybooksJson as Playbook[];

export class PlaybookStore {
  private filePath: string;
  private customPlaybooks: Playbook[];

  constructor(dataDir?: string) {
    const dir = dataDir || path.join(os.homedir(), '.contract-redliner');
    fs.mkdirSync(dir, { recursive: true });
    this.filePath = path.join(dir, 'playbooks.json');
    this.customPlaybooks = this.load();
  }

  private load(): Playbook[] {
    try {
      if (fs.existsSync(this.filePath)) {
        const parsed = JSON.parse(fs.readFileSync(this.filePath, 'utf-8'));
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {
      // ignore
    }
    return [];
  }

  private save(): void {
    try {
      fs.writeFileSync(this.filePath, JSON.stringify(this.customPlaybooks, null, 2), 'utf-8');
    } catch {
      // non-fatal
    }
  }

  list(): Playbook[] {
    return [...BUILT_IN_PLAYBOOKS, ...this.customPlaybooks];
  }

  getById(id: string): Playbook | undefined {
    return this.list().find((p) => p.id === id);
  }

  upsertCustom(pb: Omit<Playbook, 'builtIn'>): Playbook {
    const clean: Playbook = { ...pb, builtIn: false };
    const idx = this.customPlaybooks.findIndex((p) => p.id === clean.id);
    if (idx >= 0) this.customPlaybooks[idx] = clean;
    else this.customPlaybooks.push(clean);
    this.save();
    return clean;
  }
}

export function buildDemoClauses(): RedlineClause[] {
  const raw = [
    {
      id: 'c-1',
      section: 'Section 8.1',
      heading: 'Limitation of Liability',
      originalText:
        'Customer shall indemnify and remain liable without limitation for any direct, indirect, incidental, or consequential damages arising out of or related to this Agreement, while Vendor aggregate liability shall not exceed one hundred dollars ($100).',
      proposedText:
        'Neither party shall be liable for any indirect, incidental, or consequential damages. Each party\'s aggregate liability arising out of or related to this Agreement shall not exceed the total fees paid or payable by Customer to Vendor in the twelve (12) months preceding the claim.',
      riskLevel: 'dealbreaker' as const,
      playbookRule: 'Limitation of Liability',
      rationale:
        'Original text exposes you to uncapped consequential damages while capping the vendor at $100. Replaced with a standard mutual 12-month fee cap.',
    },
    {
      id: 'c-2',
      section: 'Section 4.2',
      heading: 'Customer Data & AI Training Rights',
      originalText:
        'Customer grants Vendor a perpetual, irrevocable, worldwide license to use, reproduce, and train machine learning models on all Customer Data and outputs generated through the Service.',
      proposedText:
        'Customer retains all right, title, and interest in and to Customer Data and outputs. Vendor may process Customer Data solely to provide and maintain the Service during the Term and shall not use Customer Data to train artificial intelligence or machine learning models.',
      riskLevel: 'dealbreaker' as const,
      playbookRule: 'IP & Customer Data',
      rationale:
        'Original clause grants a perpetual license to train AI models on your proprietary data. Narrowed strictly to delivering the service.',
    },
    {
      id: 'c-3',
      section: 'Section 3.3',
      heading: 'Auto-Renewal & Price Increases',
      originalText:
        'This Agreement shall automatically renew for successive three-year terms unless Customer provides written notice of non-renewal at least ninety (90) days prior to expiration, and Vendor may increase fees by up to twenty-five percent (25%) upon renewal.',
      proposedText:
        'This Agreement shall renew for successive one-year terms unless either party provides written notice of non-renewal at least thirty (30) days prior to expiration. Any renewal fee increase shall not exceed five percent (5%) over the prior term.',
      riskLevel: 'negotiate' as const,
      playbookRule: 'Auto-Renewal & Notice',
      rationale:
        'Locks you into 3-year auto-renewals with a 90-day trap and 25% price hikes. Shortened to 1-year terms, 30-day notice, and 5% uplift cap.',
    },
    {
      id: 'c-4',
      section: 'Section 5.1',
      heading: 'Payment Terms & Suspension',
      originalText:
        'All invoices are due within seven (7) days of invoice date. Vendor may immediately suspend access to the Service without notice if any amount remains unpaid.',
      proposedText:
        'All undisputed invoices are due within thirty (30) days of receipt. Vendor may suspend access only after providing fifteen (15) days written notice of overdue undisputed amounts and an opportunity to cure.',
      riskLevel: 'negotiate' as const,
      playbookRule: 'Payment Terms',
      rationale:
        'Net 7 with instant no-notice suspension is risky for production systems. Moved to Net 30 with a 15-day written cure period.',
    },
    {
      id: 'c-5',
      section: 'Section 11.2',
      heading: 'Governing Law & Venue',
      originalText:
        'This Agreement is governed by the laws of the State of Delaware, without regard to conflict of laws principles.',
      proposedText:
        'This Agreement is governed by the laws of the State of Delaware, without regard to conflict of laws principles.',
      riskLevel: 'standard' as const,
      playbookRule: 'Standard Boilerplate',
      rationale:
        'Delaware governing law is market-standard for US SaaS agreements. No change needed.',
    },
  ];

  return raw.map((c) => ({
    ...c,
    decision: c.riskLevel === 'standard' ? 'rejected' : 'pending',
    diffTokens: computeWordDiff(c.originalText, c.proposedText),
  }));
}
