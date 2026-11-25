export type RiskLevel = 'dealbreaker' | 'negotiate' | 'standard';
export type DecisionState = 'pending' | 'accepted' | 'rejected';

export interface WordDiffToken {
  type: 'equal' | 'del' | 'ins';
  text: string;
}

export interface RedlineClause {
  id: string;
  section: string;
  heading: string;
  originalText: string;
  proposedText: string;
  riskLevel: RiskLevel;
  playbookRule: string;
  rationale: string;
  decision: DecisionState;
  /** Optional paragraph index if parsed from a real .docx */
  paragraphIndex?: number;
  diffTokens: WordDiffToken[];
}

export interface PlaybookRule {
  id: string;
  topic: string;
  severity: RiskLevel;
  position: string;
  fallback: string;
}

export interface Playbook {
  id: string;
  name: string;
  audience: string;
  description: string;
  builtIn: boolean;
  rules: PlaybookRule[];
}

export interface RedlinerSessionPayload {
  sessionId: string;
  title: string;
  counterparty: string;
  playbookId: string;
  playbookName: string;
  sourceType: 'docx' | 'pdf' | 'gdoc' | 'text' | 'demo' | 'empty';
  sourceFileName?: string;
  sourceFilePath?: string;
  googleDocUrl?: string;
  clauses: RedlineClause[];
  playbooks: Playbook[];
  summary: {
    total: number;
    dealbreakers: number;
    negotiate: number;
    standard: number;
    accepted: number;
    rejected: number;
    pending: number;
  };
  error?: string;
}

export type ExportFormat = 'docx_tracked' | 'docx_clean' | 'pdf_report' | 'markdown';

export interface ExportDocxResult {
  ok: boolean;
  format?: ExportFormat;
  mimeType?: string;
  fileName: string;
  savedPath?: string;
  base64?: string;
  acceptedCount: number;
  message: string;
}
