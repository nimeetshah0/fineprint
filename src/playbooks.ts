import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { Playbook, RedlineClause } from './types.js';
import { computeWordDiff } from './docxEngine.js';

export const BUILT_IN_PLAYBOOKS: Playbook[] = [
  {
    id: 'startup-vendor-msa',
    name: 'Startup Vendor MSA',
    audience: 'Founders & Ops',
    description: 'Protects a startup signing a SaaS vendor or enterprise supplier Master Services Agreement.',
    builtIn: true,
    rules: [
      {
        id: 'r-liability',
        topic: 'Limitation of Liability',
        severity: 'dealbreaker',
        position: 'Mutual cap at 12 months of fees paid/payable; exclude consequential damages.',
        fallback: '2x annual fees for confidentiality/data breach super-cap only.',
      },
      {
        id: 'r-autorenew',
        topic: 'Auto-Renewal & Notice',
        severity: 'negotiate',
        position: 'No silent auto-renewal, or at most 30 days advance notice with email reminder.',
        fallback: 'Annual renewal with 30-day opt-out and max 5% price uplift cap.',
      },
      {
        id: 'r-ip',
        topic: 'IP & Customer Data',
        severity: 'dealbreaker',
        position: 'Customer retains 100% ownership of all data, outputs, and pre-existing IP; no AI training on customer data.',
        fallback: 'Aggregated/anonymized telemetry only for service reliability.',
      },
      {
        id: 'r-indemnity',
        topic: 'Indemnification',
        severity: 'negotiate',
        position: 'Indemnification must be mutual (vendor covers IP infringement & data breach; customer covers misuse).',
        fallback: 'Cap customer indemnity obligations at the general liability cap.',
      },
      {
        id: 'r-payment',
        topic: 'Payment Terms',
        severity: 'negotiate',
        position: 'Net 30 from receipt of undisputed invoice; right to dispute in good faith without suspension.',
        fallback: 'Net 15 with 15-day cure period before service suspension.',
      },
    ],
  },
  {
    id: 'agency-client-sow',
    name: 'Agency Client SOW / MSA',
    audience: 'Agencies & Consultants',
    description: 'Protects agencies and freelancers on scope creep, IP transfer on payment, and payment timing.',
    builtIn: true,
    rules: [
      {
        id: 'a-ip-payment',
        topic: 'IP Transfer Upon Payment',
        severity: 'dealbreaker',
        position: 'Deliverable IP transfers to Client only upon receipt of full payment; Agency retains background tools/libraries.',
        fallback: 'License to background tools granted on a non-exclusive basis.',
      },
      {
        id: 'a-payment',
        topic: 'Payment & Late Fees',
        severity: 'dealbreaker',
        position: '50% upfront deposit, Net 15–30 on milestones; work pauses if invoices are >15 days overdue.',
        fallback: 'Net 30 with 1.5%/month late fee.',
      },
      {
        id: 'a-scope',
        topic: 'Scope & Revisions',
        severity: 'negotiate',
        position: 'Includes 2 rounds of revisions; out-of-scope requests require a written Change Order.',
        fallback: 'Billed hourly at standard rate for work beyond written deliverables.',
      },
      {
        id: 'a-termination',
        topic: 'Kill Fee / Termination',
        severity: 'negotiate',
        position: 'Client pays for all completed work + 25% kill fee on remaining committed phase if terminated for convenience.',
        fallback: '14 days written notice + payment for all work in progress.',
      },
    ],
  },
  {
    id: 'mutual-nda',
    name: 'Mutual NDA (Fast-Track)',
    audience: 'Founders & Sales',
    description: 'Standard balanced mutual confidentiality agreement without hidden non-competes or perpetual traps.',
    builtIn: true,
    rules: [
      {
        id: 'n-mutual',
        topic: 'Mutuality',
        severity: 'dealbreaker',
        position: 'Confidentiality obligations must be strictly mutual, not one-way.',
        fallback: 'None — walk away from one-way NDAs for exploratory talks.',
      },
      {
        id: 'n-term',
        topic: 'Term & Survival',
        severity: 'negotiate',
        position: '2–3 year confidentiality survival (trade secrets protected while they remain trade secrets).',
        fallback: 'Up to 5 years max; reject perpetual confidentiality for general business info.',
      },
      {
        id: 'n-nonsolicit',
        topic: 'No Hidden Non-Compete / Non-Solicit',
        severity: 'dealbreaker',
        position: 'Remove any non-compete, exclusivity, or employee non-solicitation buried in an NDA.',
        fallback: 'General job postings explicitly excluded from non-solicit.',
      },
    ],
  },
];

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
