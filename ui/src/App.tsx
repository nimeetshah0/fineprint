import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  App,
  applyDocumentTheme,
  applyHostStyleVariables,
} from '@modelcontextprotocol/ext-apps';
import type {
  DecisionState,
  ExportDocxResult,
  ExportFormat,
  RedlineClause,
  RedlinerSessionPayload,
  RiskLevel,
} from '../../src/types';

type TabMode = 'redlines' | 'playbooks';
type RiskFilter = 'all' | RiskLevel | 'pending' | 'accepted';
type DiffViewMode = 'before_after' | 'inline' | 'clean';

function smoothPhraseTokens(
  tokens: RedlineClause['diffTokens'],
  collapseHeavyRewrite: boolean
): RedlineClause['diffTokens'] {
  if (tokens.length <= 2) return tokens;

  if (collapseHeavyRewrite) {
    const equalChars = tokens
      .filter((t) => t.type === 'equal')
      .reduce((s, t) => s + t.text.trim().length, 0);
    const totalChars = tokens.reduce((s, t) => s + t.text.trim().length, 0);
    if (totalChars > 0 && equalChars / totalChars < 0.45) {
      const delText = tokens
        .filter((t) => t.type !== 'ins')
        .map((t) => t.text)
        .join('');
      const insText = tokens
        .filter((t) => t.type !== 'del')
        .map((t) => t.text)
        .join('');
      return [
        { type: 'del', text: delText },
        { type: 'ins', text: insText },
      ];
    }
  }

  const out: RedlineClause['diffTokens'] = [];
  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i];
    const prev = out[out.length - 1];
    const next = tokens[i + 1];
    const wordCount = tok.text.trim().split(/\s+/).filter(Boolean).length;
    if (
      tok.type === 'equal' &&
      wordCount <= 2 &&
      prev &&
      prev.type !== 'equal' &&
      next &&
      next.type !== 'equal'
    ) {
      out.push({ type: 'del', text: tok.text });
      out.push({ type: 'ins', text: tok.text });
    } else {
      out.push(tok);
    }
  }

  const merged: RedlineClause['diffTokens'] = [];
  let i = 0;
  while (i < out.length) {
    if (out[i].type === 'equal') {
      const prev = merged[merged.length - 1];
      if (prev && prev.type === 'equal') prev.text += out[i].text;
      else merged.push({ ...out[i] });
      i++;
    } else {
      let delBuf = '';
      let insBuf = '';
      while (i < out.length && out[i].type !== 'equal') {
        if (out[i].type === 'del') delBuf += out[i].text;
        else insBuf += out[i].text;
        i++;
      }
      if (delBuf) merged.push({ type: 'del', text: delBuf });
      if (insBuf) merged.push({ type: 'ins', text: insBuf });
    }
  }
  return merged;
}

const EMPTY_PAYLOAD: RedlinerSessionPayload = {
  sessionId: '',
  title: 'Fineprint',
  counterparty: '',
  playbookId: 'startup-vendor-msa',
  playbookName: 'Startup Vendor MSA',
  sourceType: 'empty',
  clauses: [],
  playbooks: [],
  summary: {
    total: 0,
    dealbreakers: 0,
    negotiate: 0,
    standard: 0,
    accepted: 0,
    rejected: 0,
    pending: 0,
  },
};

const RISK_META: Record<RiskLevel, { label: string; cls: string }> = {
  dealbreaker: { label: 'Dealbreaker', cls: 'badge-dealbreaker' },
  negotiate: { label: 'Negotiate', cls: 'badge-negotiate' },
  standard: { label: 'Standard', cls: 'badge-standard' },
};

export function RedlinerApp() {
  const appRef = useRef<App | null>(null);
  const [connected, setConnected] = useState(false);
  const [canMessage, setCanMessage] = useState(false);
  const [displayMode, setDisplayMode] = useState<'inline' | 'fullscreen' | 'pip'>('inline');
  const [canFullscreen, setCanFullscreen] = useState(false);
  const [payload, setPayload] = useState<RedlinerSessionPayload>(EMPTY_PAYLOAD);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [exportMenuOpen, setExportMenuOpen] = useState(false);
  const [toast, setToast] = useState('');

  const [tab, setTab] = useState<TabMode>('redlines');
  const [diffViewMode, setDiffViewMode] = useState<DiffViewMode>('before_after');
  const [riskFilter, setRiskFilter] = useState<RiskFilter>('all');
  const [editingClauseId, setEditingClauseId] = useState<string | null>(null);
  const [editDraft, setEditDraft] = useState('');

  const [newPbName, setNewPbName] = useState('');
  const [newRuleTopic, setNewRuleTopic] = useState('');
  const [newRulePos, setNewRulePos] = useState('');

  const showToast = useCallback((msg: string) => {
    setToast(msg);
    setTimeout(() => setToast((cur) => (cur === msg ? '' : cur)), 4000);
  }, []);

  useEffect(() => {
    const app = new App({ name: 'Contract Redliner', version: '0.2.0' }, {});
    appRef.current = app;

    app.ontoolinput = () => setLoading(true);
    app.ontoolresult = (result) => {
      setLoading(false);
      const sc = result?.structuredContent as unknown as RedlinerSessionPayload | undefined;
      if (sc && sc.summary) {
        setPayload(sc);
        setTab('redlines');
      }
    };
    app.onhostcontextchanged = (ctx) => {
      if (ctx.theme) applyDocumentTheme(ctx.theme);
      if (ctx.styles?.variables) applyHostStyleVariables(ctx.styles.variables);
      if (ctx.displayMode) setDisplayMode(ctx.displayMode);
      if (ctx.availableDisplayModes) {
        setCanFullscreen(ctx.availableDisplayModes.includes('fullscreen'));
      }
    };

    app
      .connect()
      .then(async () => {
        setConnected(true);
        const ctx = app.getHostContext();
        if (ctx?.theme) applyDocumentTheme(ctx.theme);
        if (ctx?.styles?.variables) applyHostStyleVariables(ctx.styles.variables);
        if (ctx?.displayMode) setDisplayMode(ctx.displayMode);

        const supportsFs = Boolean(ctx?.availableDisplayModes?.includes('fullscreen'));
        setCanFullscreen(supportsFs);

        const caps = app.getHostCapabilities() as Record<string, unknown> | undefined;
        setCanMessage(Boolean(caps?.message));

        if (supportsFs && ctx?.displayMode === 'inline') {
          try {
            const res = await app.requestDisplayMode({ mode: 'fullscreen' });
            if (res?.mode) setDisplayMode(res.mode);
          } catch {
            // stays inline cleanly at 640px
          }
        }
      })
      .catch(() => setConnected(false));
  }, []);

  const recomputeSummary = (clauses: RedlineClause[]) => ({
    total: clauses.length,
    dealbreakers: clauses.filter((c) => c.riskLevel === 'dealbreaker').length,
    negotiate: clauses.filter((c) => c.riskLevel === 'negotiate').length,
    standard: clauses.filter((c) => c.riskLevel === 'standard').length,
    accepted: clauses.filter((c) => c.decision === 'accepted').length,
    rejected: clauses.filter((c) => c.decision === 'rejected').length,
    pending: clauses.filter((c) => c.decision === 'pending').length,
  });

  const handleLoadDemo = async () => {
    const app = appRef.current;
    if (!app || !connected) return;
    setLoading(true);
    try {
      const res = await app.callServerTool({ name: 'app_load_demo', arguments: {} });
      const sc = res?.structuredContent as unknown as RedlinerSessionPayload | undefined;
      if (sc && sc.summary) {
        setPayload(sc);
        setTab('redlines');
      }
    } finally {
      setLoading(false);
    }
  };

  const updateClause = async (
    clauseId: string,
    decision?: DecisionState,
    proposedText?: string
  ) => {
    setPayload((prev) => {
      const nextClauses = prev.clauses.map((c) => {
        if (c.id !== clauseId) return c;
        return {
          ...c,
          decision: decision ?? c.decision,
          proposedText: proposedText !== undefined ? proposedText : c.proposedText,
        };
      });
      return {
        ...prev,
        clauses: nextClauses,
        summary: recomputeSummary(nextClauses),
      };
    });

    const app = appRef.current;
    if (!app || !connected || !payload.sessionId) return;
    const args: Record<string, unknown> = { sessionId: payload.sessionId, clauseId };
    if (decision !== undefined) args.decision = decision;
    if (proposedText !== undefined) args.proposedText = proposedText;

    try {
      const res = await app.callServerTool({ name: 'app_update_decision', arguments: args });
      const sc = res?.structuredContent as unknown as RedlinerSessionPayload | undefined;
      if (sc && sc.summary && Array.isArray(sc.clauses)) setPayload(sc);
    } catch {
      // optimistic state preserved
    }
  };

  const handleAcceptAllDealbreakers = async () => {
    const targets = payload.clauses.filter(
      (c) => c.riskLevel === 'dealbreaker' && c.decision !== 'accepted'
    );
    if (targets.length === 0) return;
    for (const c of targets) await updateClause(c.id, 'accepted');
    showToast(`Accepted ${targets.length} dealbreaker redline(s).`);
  };

  const handleSoftenClause = async (clause: RedlineClause) => {
    const app = appRef.current;
    if (!app || !canMessage) return;
    const prompt = [
      `Please soften the proposed redline for **${clause.section} — ${clause.heading}** in *${payload.title}* so it feels collaborative and won't stall the deal, while still protecting our core playbook position (**${clause.playbookRule}**).`,
      '',
      `**Original:** ${clause.originalText}`,
      `**Current Redline:** ${clause.proposedText}`,
      '',
      `Call \`open_redliner\` (or give me the softened text) so I can update this clause.`,
    ].join('\n');

    await app.sendMessage({
      role: 'user',
      content: [{ type: 'text', text: prompt }],
    });
    showToast(`Asked agent to soften ${clause.section}.`);
  };

  const handleExport = async (format: ExportFormat) => {
    setExportMenuOpen(false);
    const app = appRef.current;
    if (!app || !connected) return;
    setExporting(true);
    try {
      const exportArgs: Record<string, unknown> = {
        sessionId: payload.sessionId || 'stateless',
        format,
        saveToDownloads: true,
        title: payload.title,
        counterparty: payload.counterparty,
        playbookName: payload.playbookName,
        clauses: payload.clauses.map((c) => ({
          id: c.id,
          section: c.section,
          heading: c.heading,
          originalText: c.originalText,
          proposedText: c.proposedText,
          riskLevel: c.riskLevel,
          playbookRule: c.playbookRule,
          rationale: c.rationale,
          decision: c.decision,
          ...(c.paragraphIndex !== undefined ? { paragraphIndex: c.paragraphIndex } : {}),
        })),
        decisions: payload.clauses.map((c) => ({
          id: c.id,
          decision: c.decision,
          proposedText: c.proposedText,
        })),
      };
      if (payload.sourceFilePath) exportArgs.sourceFilePath = payload.sourceFilePath;

      const res = await app.callServerTool({
        name: 'app_export_docx',
        arguments: exportArgs,
      });
      const out = res?.structuredContent as unknown as ExportDocxResult | undefined;
      if (out?.ok && out.base64) {
        try {
          const bin = atob(out.base64);
          const bytes = new Uint8Array(bin.length);
          for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
          const blob = new Blob([bytes], {
            type:
              out.mimeType ||
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          });
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = out.fileName;
          a.click();
          setTimeout(() => URL.revokeObjectURL(url), 1000);
        } catch {
          // saved to ~/Downloads
        }
        showToast(out.message);
      } else {
        showToast(out?.message || 'Export failed.');
      }
    } finally {
      setExporting(false);
    }
  };

  const handleCopySummaryEmail = async () => {
    const accepted = payload.clauses.filter((c) => c.decision === 'accepted');
    const lines = [
      `Hi ${payload.counterparty || 'team'},`,
      '',
      `Thanks for sending over the ${payload.title}. We reviewed it and attached a clean redline with ${accepted.length} targeted edit(s) to align with our standard terms:`,
      '',
      ...accepted.map((c) => `- ${c.section} (${c.heading}): ${c.rationale}`),
      '',
      'Everything else looks good on our end—happy to jump on a quick 10-minute call if helpful to get this wrapped up.',
    ];
    try {
      await navigator.clipboard.writeText(lines.join('\n'));
      showToast('Copied counterparty summary email.');
    } catch {
      showToast('Clipboard unavailable in this frame.');
    }
  };

  const handleSaveCustomPlaybook = async () => {
    const app = appRef.current;
    if (!app || !connected || !newPbName.trim() || !newRuleTopic.trim() || !newRulePos.trim()) {
      return;
    }
    const id = 'custom-' + newPbName.toLowerCase().replace(/[^a-z0-9]+/g, '-');
    const res = await app.callServerTool({
      name: 'app_save_playbook',
      arguments: {
        id,
        name: newPbName.trim(),
        audience: 'Custom Company Playbook',
        description: 'Saved custom redlining rules.',
        rules: [
          {
            id: 'r-1',
            topic: newRuleTopic.trim(),
            severity: 'dealbreaker',
            position: newRulePos.trim(),
            fallback: 'Escalate to founder/legal.',
          },
        ],
      },
    });
    const sc = res?.structuredContent as { playbooks?: RedlinerSessionPayload['playbooks'] } | undefined;
    if (sc?.playbooks) {
      setPayload((p) => ({ ...p, playbooks: sc.playbooks! }));
      setNewPbName('');
      setNewRuleTopic('');
      setNewRulePos('');
      showToast('Saved custom playbook!');
    }
  };

  const visibleClauses = useMemo(() => {
    return payload.clauses.filter((c) => {
      if (riskFilter === 'all') return true;
      if (riskFilter === 'pending') return c.decision === 'pending';
      if (riskFilter === 'accepted') return c.decision === 'accepted';
      return c.riskLevel === riskFilter;
    });
  }, [payload.clauses, riskFilter]);

  return (
    <div className={`shell ${canFullscreen && displayMode === 'inline' ? 'mode-inline' : ''}`}>
      {/* Material 3 Top App Bar */}
      <header className="appbar">
        <div className="appbar-top">
          <div className="appbar-title-group">
            <div className="appbar-title-row">
              <h1 className="appbar-title">
                {payload.clauses.length > 0 ? payload.title : 'Fineprint'}
              </h1>
              <span className="badge badge-privacy">Zero Retention</span>
            </div>
            {payload.clauses.length > 0 && (
              <div className="appbar-sub">
                {payload.counterparty ? `${payload.counterparty} · ` : ''}
                {payload.playbookName}
                {payload.sourceFileName ? ` · ${payload.sourceFileName}` : ''}
              </div>
            )}
          </div>

          <div className="action-group">
            <div className="segmented">
              <button
                className={`seg-btn ${tab === 'redlines' ? 'active' : ''}`}
                onClick={() => setTab('redlines')}
              >
                Sections ({payload.summary.total})
              </button>
              <button
                className={`seg-btn ${tab === 'playbooks' ? 'active' : ''}`}
                onClick={() => setTab('playbooks')}
              >
                Playbooks ({payload.playbooks.length})
              </button>
            </div>

            {canFullscreen && (
              <button
                className="m3-btn m3-btn-outlined"
                style={{ height: 30, padding: '0 12px' }}
                onClick={async () => {
                  const app = appRef.current;
                  if (!app) return;
                  const next = displayMode === 'fullscreen' ? 'inline' : 'fullscreen';
                  try {
                    const res = await app.requestDisplayMode({ mode: next });
                    if (res?.mode) setDisplayMode(res.mode);
                  } catch {
                    // ignore
                  }
                }}
              >
                {displayMode === 'fullscreen' ? 'Collapse' : 'Expand ⤢'}
              </button>
            )}
          </div>
        </div>

        {payload.clauses.length > 0 && tab === 'redlines' && (
          <div className="appbar-controls">
            <div className="chip-row">
              {(
                [
                  ['all', `All (${payload.summary.total})`],
                  ['dealbreaker', `Dealbreakers (${payload.summary.dealbreakers})`],
                  ['negotiate', `Negotiate (${payload.summary.negotiate})`],
                  ['pending', `Pending (${payload.summary.pending})`],
                  ['accepted', `Accepted (${payload.summary.accepted})`],
                ] as Array<[RiskFilter, string]>
              ).map(([id, label]) => (
                <button
                  key={id}
                  className={`m3-chip ${riskFilter === id ? 'active' : ''}`}
                  onClick={() => setRiskFilter(id)}
                >
                  {label}
                </button>
              ))}
            </div>

            <div className="segmented">
              <button
                className={`seg-btn ${diffViewMode === 'before_after' ? 'active' : ''}`}
                onClick={() => setDiffViewMode('before_after')}
              >
                Before / After
              </button>
              <button
                className={`seg-btn ${diffViewMode === 'inline' ? 'active' : ''}`}
                onClick={() => setDiffViewMode('inline')}
              >
                Inline
              </button>
              <button
                className={`seg-btn ${diffViewMode === 'clean' ? 'active' : ''}`}
                onClick={() => setDiffViewMode('clean')}
              >
                Clean
              </button>
            </div>
          </div>
        )}
      </header>

      {/* Main Feed — Each Section is a Material 3 Elevated Card */}
      <main className="main">
        {tab === 'redlines' && payload.clauses.length === 0 && !loading && (
          <div className="empty-card">
            <h2>Fineprint</h2>
            <p>
              Paste an NDA, MSA, or SOW in chat—or point your agent at a local{' '}
              <code>.docx</code> or <code>.pdf</code> file—to review each section as an
              interactive card and export native Microsoft Word Track Changes.
            </p>
            <button className="m3-btn m3-btn-filled" onClick={handleLoadDemo}>
              Load Sample SaaS Vendor MSA
            </button>
          </div>
        )}

        {tab === 'redlines' &&
          visibleClauses.map((clause) => {
            const rm = RISK_META[clause.riskLevel];
            const isEditing = editingClauseId === clause.id;
            const hasDiff = clause.originalText.trim() !== clause.proposedText.trim();

            return (
              <article
                key={clause.id}
                className={`section-card dec-${clause.decision}`}
              >
                {/* Card Header */}
                <div className="card-header">
                  <div className="card-title-wrap">
                    <div className="card-title-meta">
                      <span className="badge badge-section">{clause.section}</span>
                      <span className={`badge ${rm.cls}`}>{rm.label}</span>
                      <span className="badge badge-tonal">{clause.playbookRule}</span>
                    </div>
                    <h2 className="card-heading">{clause.heading}</h2>
                  </div>

                  {clause.decision === 'accepted' && (
                    <span className="badge badge-standard">✓ Accepted</span>
                  )}
                  {clause.decision === 'rejected' && hasDiff && (
                    <span className="badge badge-tonal">Kept Original</span>
                  )}
                </div>

                {/* Plain-English Insight */}
                <div className="insight-banner">
                  <strong>Why this matters:</strong> {clause.rationale}
                </div>

                {/* Diff / Editor Body */}
                {isEditing ? (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                    <textarea
                      className="edit-area"
                      value={editDraft}
                      onChange={(e) => setEditDraft(e.target.value)}
                    />
                    <div className="action-group">
                      <button
                        className="m3-btn m3-btn-filled"
                        onClick={async () => {
                          await updateClause(clause.id, 'accepted', editDraft);
                          setEditingClauseId(null);
                          showToast(`Saved custom wording for ${clause.section}.`);
                        }}
                      >
                        Save &amp; Accept
                      </button>
                      <button
                        className="m3-btn m3-btn-outlined"
                        onClick={() => setEditingClauseId(null)}
                      >
                        Cancel
                      </button>
                    </div>
                  </div>
                ) : !hasDiff || clause.decision === 'rejected' ? (
                  <div className="inset-pane">
                    <div className="pane-label">Original (kept as-is)</div>
                    <div className="pane-text">{clause.originalText}</div>
                  </div>
                ) : diffViewMode === 'clean' ? (
                  <div className="inset-pane inset-after">
                    <div className="pane-label">
                      {clause.decision === 'accepted'
                        ? 'Final Wording (Accepted)'
                        : 'Proposed Final Wording'}
                    </div>
                    <div className="pane-text">{clause.proposedText}</div>
                  </div>
                ) : diffViewMode === 'before_after' ? (
                  (() => {
                    const smoothed = smoothPhraseTokens(clause.diffTokens, false);
                    return (
                      <div className="compare-stack">
                        <div className="inset-pane inset-before">
                          <div className="pane-label">Original</div>
                          <div className="pane-text">
                            {smoothed
                              .filter((t) => t.type !== 'ins')
                              .map((tok, i) =>
                                tok.type === 'del' ? (
                                  <mark key={i} className="hl-del">
                                    {tok.text}
                                  </mark>
                                ) : (
                                  <span key={i}>{tok.text}</span>
                                )
                              )}
                          </div>
                        </div>

                        <div className="inset-pane inset-after">
                          <div className="pane-label">Proposed Redline</div>
                          <div className="pane-text">
                            {smoothed
                              .filter((t) => t.type !== 'del')
                              .map((tok, i) =>
                                tok.type === 'ins' ? (
                                  <mark key={i} className="hl-ins">
                                    {tok.text}
                                  </mark>
                                ) : (
                                  <span key={i}>{tok.text}</span>
                                )
                              )}
                          </div>
                        </div>
                      </div>
                    );
                  })()
                ) : (
                  <div className="inset-pane">
                    <div className="pane-text">
                      {smoothPhraseTokens(clause.diffTokens, true).map((tok, i) => {
                        if (tok.type === 'del') {
                          return (
                            <del key={i} className="tok-del">
                              {tok.text}
                            </del>
                          );
                        }
                        if (tok.type === 'ins') {
                          return (
                            <ins key={i} className="tok-ins">
                              {tok.text}
                            </ins>
                          );
                        }
                        return <span key={i}>{tok.text}</span>;
                      })}
                    </div>
                  </div>
                )}

                {/* Card Actions */}
                <div className="card-actions">
                  <div className="action-group">
                    {hasDiff && (
                      <button
                        className={`m3-btn ${
                          clause.decision === 'accepted'
                            ? 'm3-btn-tonal-success'
                            : 'm3-btn-tonal'
                        }`}
                        onClick={() =>
                          updateClause(
                            clause.id,
                            clause.decision === 'accepted' ? 'pending' : 'accepted'
                          )
                        }
                      >
                        {clause.decision === 'accepted' ? '✓ Accepted' : 'Accept Redline'}
                      </button>
                    )}

                    <button
                      className={`m3-btn ${
                        clause.decision === 'rejected'
                          ? 'm3-btn-tonal-error'
                          : 'm3-btn-outlined'
                      }`}
                      onClick={() =>
                        updateClause(
                          clause.id,
                          clause.decision === 'rejected' ? 'pending' : 'rejected'
                        )
                      }
                    >
                      {clause.decision === 'rejected' ? 'Keeping Original' : 'Keep Original'}
                    </button>

                    <button
                      className="m3-btn m3-btn-text"
                      onClick={() => {
                        setEditingClauseId(clause.id);
                        setEditDraft(clause.proposedText);
                      }}
                    >
                      Edit
                    </button>
                  </div>

                  {canMessage && hasDiff && (
                    <button
                      className="m3-btn m3-btn-text"
                      onClick={() => handleSoftenClause(clause)}
                    >
                      Soften with AI
                    </button>
                  )}
                </div>
              </article>
            );
          })}

        {tab === 'playbooks' && (
          <>
            {payload.playbooks.map((pb) => (
              <section key={pb.id} className="section-card">
                <div className="card-header">
                  <div>
                    <h2 className="card-heading">{pb.name}</h2>
                    <div className="appbar-sub">
                      {pb.audience} · {pb.description}
                    </div>
                  </div>
                  <span className="badge badge-tonal">
                    {pb.builtIn ? 'Built-in' : 'Custom'}
                  </span>
                </div>
                <div className="pb-rule-list">
                  {pb.rules.map((r) => (
                    <div key={r.id} className="pb-rule">
                      <div
                        style={{
                          display: 'flex',
                          justifyContent: 'space-between',
                          alignItems: 'center',
                          marginBottom: 4,
                        }}
                      >
                        <strong>{r.topic}</strong>
                        <span className={`badge ${RISK_META[r.severity].cls}`}>
                          {RISK_META[r.severity].label}
                        </span>
                      </div>
                      <div>{r.position}</div>
                      <div className="appbar-sub" style={{ marginTop: 2 }}>
                        Fallback: {r.fallback}
                      </div>
                    </div>
                  ))}
                </div>
              </section>
            ))}

            <section className="section-card">
              <div className="card-header">
                <div>
                  <h2 className="card-heading">Add or Edit Community Playbooks</h2>
                  <div className="appbar-sub">
                    Built-in playbooks live in <code>src/playbooks.json</code>. Open a GitHub PR to add a new playbook or refine existing rules for everyone.
                  </div>
                </div>
                <button
                  className="m3-btn m3-btn-tonal"
                  onClick={() => {
                    const url = 'https://github.com/nimeetshah0/fineprint/edit/main/src/playbooks.json';
                    const app = appRef.current;
                    if (app && connected) {
                      void app.openLink({ url });
                    } else {
                      window.open(url, '_blank', 'noopener,noreferrer');
                    }
                  }}
                >
                  Open GitHub PR ↗
                </button>
              </div>
            </section>

            <section className="section-card">
              <h2 className="card-heading">Add Local Session Playbook</h2>
              <input
                className="m3-input"
                placeholder="Playbook name (e.g. Acme Enterprise Sales MSA)"
                value={newPbName}
                onChange={(e) => setNewPbName(e.target.value)}
              />
              <input
                className="m3-input"
                placeholder="Rule topic (e.g. Liability Cap)"
                value={newRuleTopic}
                onChange={(e) => setNewRuleTopic(e.target.value)}
              />
              <input
                className="m3-input"
                placeholder="Your position (e.g. Cap at 12 months fees; no consequential damages)"
                value={newRulePos}
                onChange={(e) => setNewRulePos(e.target.value)}
              />
              <div>
                <button
                  className="m3-btn m3-btn-filled"
                  disabled={!newPbName.trim() || !newRuleTopic.trim() || !newRulePos.trim()}
                  onClick={handleSaveCustomPlaybook}
                >
                  Save Playbook
                </button>
              </div>
            </section>
          </>
        )}
      </main>

      {/* Material 3 Bottom Action Bar */}
      {payload.clauses.length > 0 && (
        <footer className="bottombar">
          <div className="action-group">
            <span className="appbar-sub">
              {payload.summary.accepted} accepted · {payload.summary.pending} pending ·{' '}
              <em>Drafting aid only, not legal advice</em>
            </span>
            {toast && <span className="toast">{toast}</span>}
          </div>

          <div className="action-group">
            {payload.summary.dealbreakers > 0 && (
              <button className="m3-btn m3-btn-tonal" onClick={handleAcceptAllDealbreakers}>
                Accept Dealbreakers
              </button>
            )}
            {payload.summary.accepted > 0 && (
              <button className="m3-btn m3-btn-outlined" onClick={handleCopySummaryEmail}>
                Copy Email
              </button>
            )}
            <div className="export-menu-wrap">
              <button
                className="m3-btn m3-btn-filled"
                disabled={exporting}
                onClick={() => setExportMenuOpen((o) => !o)}
              >
                {exporting ? 'Exporting…' : `Export (${payload.summary.accepted}) ▾`}
              </button>

              {exportMenuOpen && (
                <div className="export-menu">
                  <button
                    className="export-item"
                    onClick={() => handleExport('docx_tracked')}
                  >
                    <span className="export-item-title">Word — Track Changes (.docx)</span>
                    <span className="export-item-sub">
                      Native Word redlines + margin comments
                    </span>
                  </button>
                  <button
                    className="export-item"
                    onClick={() => handleExport('docx_clean')}
                  >
                    <span className="export-item-title">Word — Clean Final (.docx)</span>
                    <span className="export-item-sub">
                      Accepted edits applied, ready to sign
                    </span>
                  </button>
                  <button
                    className="export-item"
                    onClick={() => handleExport('pdf_report')}
                  >
                    <span className="export-item-title">PDF — Redline Report (.pdf)</span>
                    <span className="export-item-sub">
                      Shareable executive summary &amp; edits
                    </span>
                  </button>
                  <button
                    className="export-item"
                    onClick={() => handleExport('markdown')}
                  >
                    <span className="export-item-title">Markdown — Memo (.md)</span>
                    <span className="export-item-sub">
                      For Notion, Slack, or email
                    </span>
                  </button>
                </div>
              )}
            </div>
          </div>
        </footer>
      )}
    </div>
  );
}
