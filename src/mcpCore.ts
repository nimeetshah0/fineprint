import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import {
  computeWordDiff,
  exportTrackedChangesDocx,
  type ParsedDocx,
} from './docxEngine.js';
import { exportMarkdownMemo, exportRedlinePdfReport } from './pdfExport.js';
import { BUILT_IN_PLAYBOOKS, buildDemoClauses } from './playbooks.js';
import type {
  ExportDocxResult,
  ExportFormat,
  Playbook,
  RedlineClause,
  RedlinerSessionPayload,
} from './types.js';

export const APP_URI = 'ui://contract-redliner/app.html';
export const APP_MIME = 'text/html;profile=mcp-app';

export interface InMemorySession {
  sessionId: string;
  title: string;
  counterparty: string;
  playbookId: string;
  sourceType: 'docx' | 'pdf' | 'text' | 'demo' | 'empty';
  sourceFileName?: string;
  sourceFilePath?: string;
  sourceDocx?: ParsedDocx;
  clauses: RedlineClause[];
}

export interface McpCoreOptions {
  readAppHtml: () => string;
  /** When true (Cloudflare Worker / public HTTP), disables all server filesystem access and cross-user playbook writes */
  publicMode: boolean;
  listPlaybooks: () => Playbook[];
  getPlaybookById: (id: string) => Playbook | undefined;
  saveCustomPlaybook?: (pb: Omit<Playbook, 'builtIn'>) => void;
  sessions?: Map<string, InMemorySession>;
  readLocalContract?: (filePath: string) => Promise<{
    content: Array<{ type: 'text'; text: string }>;
    structuredContent: Record<string, unknown>;
  }>;
  rehydrateLocalDocx?: (filePath: string) => Promise<ParsedDocx | undefined>;
  saveToDownloads?: (fileName: string, buffer: Buffer) => string | undefined;
}

function randomId(prefix: string): string {
  const bytes = new Uint8Array(6);
  crypto.getRandomValues(bytes);
  const hex = Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
  return `${prefix}_${hex}`;
}

function summarizeClauses(clauses: RedlineClause[]) {
  return {
    total: clauses.length,
    dealbreakers: clauses.filter((c) => c.riskLevel === 'dealbreaker').length,
    negotiate: clauses.filter((c) => c.riskLevel === 'negotiate').length,
    standard: clauses.filter((c) => c.riskLevel === 'standard').length,
    accepted: clauses.filter((c) => c.decision === 'accepted').length,
    rejected: clauses.filter((c) => c.decision === 'rejected').length,
    pending: clauses.filter((c) => c.decision === 'pending').length,
  };
}

const ClauseInputSchema = z.object({
  id: z.string().optional(),
  section: z.string().optional().describe('Clause or section number, e.g. "Section 8.1"'),
  heading: z.string().describe('Short clause title, e.g. "Limitation of Liability"'),
  originalText: z.string().describe('Exact original contract wording'),
  proposedText: z
    .string()
    .describe('Proposed redlined wording (same as originalText if standard/no change)'),
  riskLevel: z.enum(['dealbreaker', 'negotiate', 'standard']),
  playbookRule: z.string().optional().describe('Which playbook rule triggered this flag'),
  rationale: z.string().describe('Plain-English 1–2 sentence explanation of why this matters'),
  paragraphIndex: z
    .number()
    .optional()
    .describe('Paragraph index from read_contract, if applicable'),
});

export function createContractRedlinerServer(opts: McpCoreOptions): McpServer {
  const sessions = opts.sessions || new Map<string, InMemorySession>();

  const toPayload = (session: InMemorySession, error?: string): RedlinerSessionPayload => {
    const allPbs = opts.listPlaybooks();
    const pb = opts.getPlaybookById(session.playbookId) || allPbs[0] || BUILT_IN_PLAYBOOKS[0];
    return {
      sessionId: session.sessionId,
      title: session.title,
      counterparty: session.counterparty,
      playbookId: pb.id,
      playbookName: pb.name,
      sourceType: session.sourceType,
      sourceFileName: session.sourceFileName || session.sourceDocx?.fileName,
      sourceFilePath: opts.publicMode ? undefined : session.sourceFilePath,
      clauses: session.clauses,
      playbooks: allPbs,
      summary: summarizeClauses(session.clauses),
      error,
    };
  };

  const server = new McpServer(
    {
      name: 'contract-redliner',
      version: '0.3.0',
    },
    {
      capabilities: {
        resources: {},
        tools: {},
      },
    }
  );

  // 1. UI HTML resource
  server.registerResource(
    'contract-redliner-app',
    APP_URI,
    {
      title: 'Contract Redliner UI',
      description: 'Interactive contract redline workbench with native Word Track Changes export',
      mimeType: APP_MIME,
      _meta: {
        ui: {
          prefersBorder: false,
        },
      },
    },
    async () => ({
      contents: [
        {
          uri: APP_URI,
          mimeType: APP_MIME,
          text: opts.readAppHtml(),
        },
      ],
    })
  );

  // 2. List playbooks
  server.registerTool(
    'list_playbooks',
    {
      title: 'List Contract Playbooks',
      description:
        'Returns built-in (and local custom) contract review playbooks (Startup Vendor MSA, Agency Client SOW, Mutual NDA) with risk rules and fallback positions. Call this before reviewing a contract so you redline against the right playbook.',
      inputSchema: {},
    },
    async () => {
      const playbooks = opts.listPlaybooks();
      const text = playbooks
        .map(
          (pb) =>
            `### ${pb.name} (\`${pb.id}\`) — ${pb.audience}\n${pb.description}\n` +
            pb.rules
              .map(
                (r) =>
                  `- **${r.topic}** [${r.severity.toUpperCase()}]: ${r.position} *(Fallback: ${r.fallback})*`
              )
              .join('\n')
        )
        .join('\n\n');

      return {
        content: [{ type: 'text', text }],
        structuredContent: { playbooks },
      };
    }
  );

  // 3. Local-only file readers (disabled in public HTTP / Cloudflare Worker mode!)
  if (!opts.publicMode && opts.readLocalContract) {
    const reader = opts.readLocalContract;
    server.registerTool(
      'read_contract',
      {
        title: 'Read Contract (.docx or .pdf, Zero Retention)',
        description:
          'Reads a local Microsoft Word (.docx) or PDF (.pdf) contract purely in memory and returns numbered clauses so you can review them against a playbook and pass them to open_redliner with the returned sessionId.',
        inputSchema: {
          filePath: z.string().describe('Absolute path to a local .docx or .pdf contract file'),
        },
      },
      async ({ filePath }) => reader(filePath)
    );

    server.registerTool(
      'read_docx_contract',
      {
        title: 'Read .docx / .pdf Contract (Zero Retention)',
        description:
          'Reads a local Microsoft Word (.docx) or PDF (.pdf) contract in memory and returns numbered paragraphs so you can analyze risky clauses and pass them to open_redliner with the returned sessionId.',
        inputSchema: {
          filePath: z.string().describe('Absolute path to a local .docx or .pdf contract file'),
        },
      },
      async ({ filePath }) => reader(filePath)
    );
  }

  // 4. Main MCP App tool
  server.registerTool(
    'open_redliner',
    {
      title: 'Contract Redliner',
      description:
        'Opens the interactive Contract Redliner UI showing Before/After clause cards, risk badges, Accept/Keep/Soften controls, playbooks, and 1-click export to Word Track Changes (.docx), Clean Final (.docx), PDF Report, or Markdown. Call with structured redlines after reviewing a contract, or with {} to open the workbench with a sample demo contract.',
      inputSchema: {
        title: z.string().optional().describe('Contract title, e.g. "Acme SaaS MSA"'),
        counterparty: z.string().optional().describe('Counterparty name, e.g. "Acme Corp"'),
        playbookId: z
          .string()
          .optional()
          .describe('Playbook ID (e.g. startup-vendor-msa, agency-client-sow, mutual-nda)'),
        sessionId: z
          .string()
          .optional()
          .describe('Optional sessionId from read_contract to preserve original .docx formatting'),
        redlines: z
          .array(ClauseInputSchema)
          .optional()
          .describe('Structured clause redlines generated from the contract'),
      },
      _meta: {
        ui: {
          resourceUri: APP_URI,
        },
        'com.google.antigravity/ui': {
          entrypoints: [
            { type: 'aux_pane', title: 'Contract Redliner' },
            { type: 'global', title: 'Contract Redliner' },
          ],
        },
      },
    },
    async (args) => {
      const input = args || {};
      const existing = input.sessionId ? sessions.get(input.sessionId) : undefined;
      const sessionId = existing?.sessionId || randomId('sess');

      const clauses: RedlineClause[] = (input.redlines || []).map((r, i) => ({
        id: r.id || `c-${i + 1}`,
        section: r.section || `Clause ${i + 1}`,
        heading: r.heading,
        originalText: r.originalText,
        proposedText: r.proposedText,
        riskLevel: r.riskLevel,
        playbookRule: r.playbookRule || 'Playbook Rule',
        rationale: r.rationale,
        paragraphIndex: r.paragraphIndex,
        decision:
          r.originalText.trim() === r.proposedText.trim() ? 'rejected' : 'pending',
        diffTokens: computeWordDiff(r.originalText, r.proposedText),
      }));

      const session: InMemorySession = {
        sessionId,
        title: input.title || existing?.title || 'Contract Redliner',
        counterparty: input.counterparty || existing?.counterparty || '',
        playbookId: input.playbookId || existing?.playbookId || 'startup-vendor-msa',
        sourceType:
          clauses.length === 0
            ? 'empty'
            : existing?.sourceType === 'pdf'
              ? 'pdf'
              : existing?.sourceDocx
                ? 'docx'
                : 'text',
        sourceFileName: existing?.sourceFileName,
        sourceFilePath: existing?.sourceFilePath,
        sourceDocx: existing?.sourceDocx,
        clauses,
      };

      if (!opts.publicMode) {
        sessions.set(sessionId, session);
      }
      const payload = toPayload(session);

      const summaryText =
        clauses.length === 0
          ? 'Opened Contract Redliner workbench. Paste a contract in chat (or click "Load Sample SaaS Vendor MSA" in the UI) to review redlines and export Word Track Changes.'
          : [
              `### Contract Redliner: ${payload.title}`,
              `Playbook: **${payload.playbookName}** · ${payload.summary.total} clauses (${payload.summary.dealbreakers} dealbreakers, ${payload.summary.negotiate} negotiate, ${payload.summary.standard} standard)`,
              '',
              ...clauses.map(
                (c) =>
                  `- **${c.section} — ${c.heading}** [${c.riskLevel.toUpperCase()}]: ${c.rationale}`
              ),
            ].join('\n');

      return {
        content: [{ type: 'text', text: summaryText }],
        structuredContent: payload as unknown as Record<string, unknown>,
      };
    }
  );

  // 5. App-only: load sample SaaS MSA demo
  server.registerTool(
    'app_load_demo',
    {
      title: 'Load Demo Contract (App Only)',
      description: 'Loads a sample SaaS Vendor MSA with 5 classic founder traps.',
      inputSchema: {},
      _meta: { ui: { visibility: ['app'] } },
    },
    async () => {
      const sessionId = randomId('demo');
      const session: InMemorySession = {
        sessionId,
        title: 'Acme Cloud Master Services Agreement',
        counterparty: 'Acme Cloud Inc.',
        playbookId: 'startup-vendor-msa',
        sourceType: 'demo',
        clauses: buildDemoClauses(),
      };
      if (!opts.publicMode) sessions.set(sessionId, session);
      const payload = toPayload(session);
      return {
        content: [{ type: 'text', text: 'Loaded sample SaaS Vendor MSA demo.' }],
        structuredContent: payload as unknown as Record<string, unknown>,
      };
    }
  );

  // 6. App-only: update clause decision
  server.registerTool(
    'app_update_decision',
    {
      title: 'Update Clause Decision (App Only)',
      description: 'Updates accept/reject/pending state or custom proposed wording for a clause.',
      inputSchema: {
        sessionId: z.string(),
        clauseId: z.string(),
        decision: z.enum(['pending', 'accepted', 'rejected']).optional(),
        proposedText: z.string().optional(),
      },
      _meta: { ui: { visibility: ['app'] } },
    },
    async ({ sessionId, clauseId, decision, proposedText }) => {
      const session = sessions.get(sessionId);
      if (!session) {
        return {
          content: [{ type: 'text', text: 'Stateless mode: decision kept in UI.' }],
          structuredContent: { ok: true },
        };
      }
      const clause = session.clauses.find((c) => c.id === clauseId);
      if (clause) {
        if (decision) clause.decision = decision;
        if (proposedText !== undefined) {
          clause.proposedText = proposedText;
          clause.diffTokens = computeWordDiff(clause.originalText, clause.proposedText);
        }
      }
      const payload = toPayload(session);
      return {
        content: [{ type: 'text', text: `Updated ${clauseId}` }],
        structuredContent: payload as unknown as Record<string, unknown>,
      };
    }
  );

  // 7. App-only: stateless-safe 4-format exporter
  server.registerTool(
    'app_export_docx',
    {
      title: 'Export Redlined Contract (App Only)',
      description:
        'Exports accepted redlines as Word Track Changes (.docx), Word Clean Final (.docx), PDF Redline Report (.pdf), or Markdown Memo (.md).',
      inputSchema: {
        sessionId: z.string(),
        format: z
          .enum(['docx_tracked', 'docx_clean', 'pdf_report', 'markdown'])
          .optional(),
        saveToDownloads: z.boolean().optional(),
        title: z.string().optional(),
        counterparty: z.string().optional(),
        playbookName: z.string().optional(),
        sourceFilePath: z.string().optional(),
        clauses: z
          .array(
            z.object({
              id: z.string(),
              section: z.string(),
              heading: z.string(),
              originalText: z.string(),
              proposedText: z.string(),
              riskLevel: z.enum(['dealbreaker', 'negotiate', 'standard']),
              playbookRule: z.string(),
              rationale: z.string(),
              decision: z.enum(['pending', 'accepted', 'rejected']),
              paragraphIndex: z.number().optional(),
            })
          )
          .optional(),
        decisions: z
          .array(
            z.object({
              id: z.string(),
              decision: z.enum(['pending', 'accepted', 'rejected']),
              proposedText: z.string().optional(),
            })
          )
          .optional(),
      },
      _meta: { ui: { visibility: ['app'] } },
    },
    async ({
      sessionId,
      format,
      saveToDownloads,
      title,
      counterparty,
      playbookName,
      sourceFilePath,
      clauses: uiClauses,
      decisions,
    }) => {
      const exportFormat: ExportFormat = format || 'docx_tracked';
      const session = sessions.get(sessionId);

      let finalClauses: RedlineClause[] = [];
      let finalTitle = title || 'Contract';
      let finalCounterparty = counterparty || '';
      let finalPlaybookName = playbookName || 'Playbook';
      let sourceDocx = session?.sourceDocx;

      if (session) {
        finalTitle = session.title;
        finalCounterparty = session.counterparty;
        const pb = opts.getPlaybookById(session.playbookId) || opts.listPlaybooks()[0];
        finalPlaybookName = pb.name;
        finalClauses = session.clauses;

        if (decisions && decisions.length > 0) {
          const decMap = new Map(decisions.map((d) => [d.id, d]));
          for (const c of finalClauses) {
            const d = decMap.get(c.id);
            if (d) {
              c.decision = d.decision;
              if (d.proposedText !== undefined) {
                c.proposedText = d.proposedText;
                c.diffTokens = computeWordDiff(c.originalText, c.proposedText);
              }
            }
          }
        }
      } else if (uiClauses && uiClauses.length > 0) {
        finalClauses = uiClauses.map((c) => ({
          ...c,
          diffTokens: computeWordDiff(c.originalText, c.proposedText),
        }));
        if (!opts.publicMode && !sourceDocx && sourceFilePath && opts.rehydrateLocalDocx) {
          sourceDocx = await opts.rehydrateLocalDocx(sourceFilePath);
        }
      } else {
        const fail: ExportDocxResult = {
          ok: false,
          fileName: '',
          acceptedCount: 0,
          message: 'No clauses available to export.',
        };
        return {
          content: [{ type: 'text', text: fail.message }],
          structuredContent: fail as unknown as Record<string, unknown>,
        };
      }

      const safeBase = finalTitle.replace(/[^a-zA-Z0-9._-]+/g, '_').slice(0, 60) || 'contract';
      const acceptedCount = finalClauses.filter((c) => c.decision === 'accepted').length;

      let buffer: Buffer;
      let fileName: string;
      let mimeType: string;

      if (exportFormat === 'docx_clean') {
        buffer = await exportTrackedChangesDocx({
          title: finalTitle,
          counterparty: finalCounterparty,
          playbookName: finalPlaybookName,
          clauses: finalClauses,
          sourceDocx,
          cleanFinal: true,
        });
        fileName = `${safeBase}_CLEAN_FINAL.docx`;
        mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
      } else if (exportFormat === 'pdf_report') {
        buffer = exportRedlinePdfReport({
          title: finalTitle,
          counterparty: finalCounterparty,
          playbookName: finalPlaybookName,
          clauses: finalClauses,
        });
        fileName = `${safeBase}_REDLINE_REPORT.pdf`;
        mimeType = 'application/pdf';
      } else if (exportFormat === 'markdown') {
        buffer = exportMarkdownMemo({
          title: finalTitle,
          counterparty: finalCounterparty,
          playbookName: finalPlaybookName,
          clauses: finalClauses,
        });
        fileName = `${safeBase}_REDLINE_MEMO.md`;
        mimeType = 'text/markdown';
      } else {
        buffer = await exportTrackedChangesDocx({
          title: finalTitle,
          counterparty: finalCounterparty,
          playbookName: finalPlaybookName,
          clauses: finalClauses,
          sourceDocx,
          cleanFinal: false,
        });
        fileName = `${safeBase}_REDLINED.docx`;
        mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
      }

      const savedPath =
        !opts.publicMode && saveToDownloads && opts.saveToDownloads
          ? opts.saveToDownloads(fileName, buffer)
          : undefined;

      const result: ExportDocxResult = {
        ok: true,
        format: exportFormat,
        mimeType,
        fileName,
        savedPath,
        base64: buffer.toString('base64'),
        acceptedCount,
        message: savedPath ? `Saved ${fileName} to ~/Downloads` : `Exported ${fileName}`,
      };

      return {
        content: [{ type: 'text', text: result.message }],
        structuredContent: result as unknown as Record<string, unknown>,
      };
    }
  );

  // 8. App-only: save custom playbook (persisted locally; read-only notice on public Worker)
  server.registerTool(
    'app_save_playbook',
    {
      title: 'Save Custom Playbook (App Only)',
      description: 'Creates or updates a custom company contract playbook.',
      inputSchema: {
        id: z.string(),
        name: z.string(),
        audience: z.string(),
        description: z.string(),
        rules: z.array(
          z.object({
            id: z.string(),
            topic: z.string(),
            severity: z.enum(['dealbreaker', 'negotiate', 'standard']),
            position: z.string(),
            fallback: z.string(),
          })
        ),
      },
      _meta: { ui: { visibility: ['app'] } },
    },
    async (pb) => {
      if (!opts.publicMode && opts.saveCustomPlaybook) {
        opts.saveCustomPlaybook(pb);
        return {
          content: [{ type: 'text', text: `Saved playbook ${pb.name}` }],
          structuredContent: { playbooks: opts.listPlaybooks() },
        };
      }
      // On public stateless Worker, return built-in + this session's custom playbook without writing globally
      return {
        content: [{ type: 'text', text: `Added ${pb.name} for this session.` }],
        structuredContent: {
          playbooks: [...opts.listPlaybooks(), { ...pb, builtIn: false }],
        },
      };
    }
  );

  return server;
}
