import { WebStandardStreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js';
import { createContractRedlinerServer } from './mcpCore.js';
import { BUILT_IN_PLAYBOOKS } from './playbooks.js';
// Bundled as static text by Wrangler via [[rules]] type = "Text"
// @ts-ignore — Wrangler bundles .html as string
import APP_HTML from '../dist/ui/index.html';

const CORS_HEADERS: Record<string, string> = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Accept, Mcp-Session-Id, Authorization',
  'Access-Control-Expose-Headers': 'Mcp-Session-Id',
};

const ICON_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128" fill="none">
  <defs>
    <linearGradient id="fp-bg" x1="12" y1="8" x2="116" y2="120" gradientUnits="userSpaceOnUse">
      <stop offset="0%" stop-color="#1E1E24"/>
      <stop offset="100%" stop-color="#0D0D10"/>
    </linearGradient>
    <linearGradient id="fp-red" x1="34" y1="62" x2="94" y2="62" gradientUnits="userSpaceOnUse">
      <stop offset="0%" stop-color="#FF5A5F"/>
      <stop offset="100%" stop-color="#E11D48"/>
    </linearGradient>
  </defs>
  <rect width="128" height="128" rx="30" fill="url(#fp-bg)"/>
  <rect x="2" y="2" width="124" height="124" rx="28" stroke="rgba(255,255,255,0.08)" stroke-width="2"/>
  <!-- Stylized geometric F with editorial redline middle stroke -->
  <path d="M38 30C38 26.6863 40.6863 24 44 24H88C91.3137 24 94 26.6863 94 30V36C94 39.3137 91.3137 42 88 42H56V98C56 101.314 53.3137 104 50 104H44C40.6863 104 38 101.314 38 98V30Z" fill="#F8FAFC"/>
  <!-- Bold redline accent stroke (with subtle tilted editorial cut) -->
  <path d="M36 60C36 56.6863 38.6863 54 42 54H84C87.3137 54 90 56.6863 90 60V64C90 67.3137 87.3137 70 84 70H42C38.6863 70 36 67.3137 36 64V60Z" fill="url(#fp-red)"/>
  <!-- Subtle fineprint dot accent -->
  <circle cx="86" cy="94" r="9" fill="#34D399"/>
</svg>`;

function landingHtml(origin: string): string {
  const mcpUrl = `${origin}/mcp`;
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Fineprint — Free Contract Redliner MCP App</title>
  <link rel="icon" type="image/svg+xml" href="/favicon.svg" />
  <style>
    :root { color-scheme: dark; }
    body { margin: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Inter, sans-serif; background: #111318; color: #e2e2e9; display: grid; place-items: center; min-height: 100vh; padding: 24px; }
    .card { max-width: 620px; width: 100%; background: #191c22; border: 1px solid rgba(255,255,255,0.1); border-radius: 20px; padding: 28px 30px; box-shadow: 0 10px 30px rgba(0,0,0,0.35); }
    .badge { display: inline-block; font-size: 11px; font-weight: 600; padding: 4px 10px; border-radius: 999px; background: rgba(129,201,149,0.16); color: #c8f5d4; margin-bottom: 10px; }
    h1 { margin: 0 0 8px; font-size: 24px; letter-spacing: -0.02em; }
    p { margin: 0 0 18px; color: #a8aebc; line-height: 1.55; font-size: 14.5px; }
    .url-box { display: flex; gap: 8px; background: #111318; border: 1px solid rgba(255,255,255,0.14); border-radius: 12px; padding: 10px 12px; align-items: center; margin-bottom: 18px; }
    code { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 13px; color: #8ab4f8; word-break: break-all; flex: 1; }
    button { border: none; background: #8ab4f8; color: #0b2545; font-weight: 600; padding: 7px 14px; border-radius: 999px; cursor: pointer; font-size: 12.5px; }
    ul { margin: 0 0 20px; padding-left: 18px; color: #a8aebc; font-size: 13.5px; line-height: 1.7; }
    strong { color: #e2e2e9; }
    .footer-links { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding-top: 16px; border-top: 1px solid rgba(255,255,255,0.08); flex-wrap: wrap; }
    .gh-link { display: inline-flex; align-items: center; gap: 7px; color: #e2e2e9; text-decoration: none; font-size: 13px; font-weight: 600; background: #232730; border: 1px solid rgba(255,255,255,0.14); padding: 7px 14px; border-radius: 999px; }
    .gh-link:hover { border-color: #8ab4f8; color: #8ab4f8; }
    .meta-note { font-size: 12px; color: #747b8a; }
  </style>
</head>
<body>
  <main class="card">
    <span class="badge">Free &amp; Open Source · Zero Contract Retention · MCP Apps + Skills</span>
    <h1>Fineprint</h1>
    <p>Catch hidden contract traps before you sign. An interactive Material 3 contract redliner for <strong>Claude</strong>, <strong>ChatGPT</strong>, and <strong>Google Antigravity</strong>—supporting <strong>Google Docs</strong>, drag-and-drop <code>.docx</code> / <code>.pdf</code> uploads, and native Microsoft Word Track Changes (<code>.docx</code>) with margin comments.</p>
    <div class="url-box">
      <code id="u">${mcpUrl}</code>
      <button onclick="navigator.clipboard.writeText('${mcpUrl}');this.textContent='Copied!'">Copy URL</button>
    </div>
    <ul>
      <li><strong>Claude.ai / Mobile:</strong> Settings → Connectors → Add Custom Connector → paste URL above.</li>
      <li><strong>ChatGPT:</strong> Settings → Connectors / Apps → Add MCP Server → paste URL above.</li>
      <li><strong>Google Docs &amp; Drag-and-Drop:</strong> Paste any link-shared Google Docs URL or drop a <code>.docx</code> / <code>.pdf</code> directly into the Fineprint UI—exported Track Changes preserve formatting and re-import into Google Docs as native Suggestions.</li>
      <li><strong>Playbooks via Skills over MCP + GitHub PRs:</strong> Built-in playbooks are delivered dynamically via <strong>Skills over MCP</strong> (<code>skill://fineprint/SKILL.md</code>) and sourced from <a href="https://github.com/nimeetshah0/fineprint/blob/main/src/playbooks.json" target="_blank" rel="noopener noreferrer" style="color:#8ab4f8;"><code>src/playbooks.json</code></a>—open a <a href="https://github.com/nimeetshah0/fineprint/pulls" target="_blank" rel="noopener noreferrer" style="color:#8ab4f8;">GitHub PR</a> to add or refine community playbooks for everyone.</li>
    </ul>
    <div class="footer-links">
      <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
        <a class="gh-link" href="https://github.com/nimeetshah0/fineprint" target="_blank" rel="noopener noreferrer">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M12 .5C5.65.5.5 5.65.5 12c0 5.08 3.29 9.39 7.86 10.91.58.11.79-.25.79-.56 0-.28-.01-1.02-.02-2-3.2.69-3.88-1.54-3.88-1.54-.52-1.33-1.28-1.69-1.28-1.69-1.05-.71.08-.7.08-.7 1.16.08 1.77 1.19 1.77 1.19 1.03 1.76 2.7 1.25 3.36.96.1-.75.4-1.25.73-1.54-2.55-.29-5.24-1.28-5.24-5.68 0-1.25.45-2.28 1.18-3.08-.12-.29-.51-1.46.11-3.04 0 0 .97-.31 3.17 1.18a11 11 0 0 1 5.77 0c2.2-1.49 3.17-1.18 3.17-1.18.62 1.58.23 2.75.11 3.04.74.8 1.18 1.83 1.18 3.08 0 4.41-2.69 5.39-5.25 5.67.41.36.78 1.06.78 2.13 0 1.54-.01 2.78-.01 3.16 0 .31.21.68.8.56A11.51 11.51 0 0 0 23.5 12C23.5 5.65 18.35.5 12 .5z"/></svg>
          View on GitHub
        </a>
        <a class="gh-link" href="https://github.com/nimeetshah0/fineprint/pulls" target="_blank" rel="noopener noreferrer">Submit Playbook PR</a>
        <a class="gh-link" href="/privacy">Privacy &amp; Disclaimer</a>
      </div>
      <span class="meta-note">Drafting aid only · Not legal advice · MIT</span>
    </div>
  </main>
</body>
</html>`;
}

function privacyHtml(): string {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Privacy Policy &amp; Disclaimer — Fineprint</title>
  <style>
    :root { color-scheme: dark; }
    body { margin: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Inter, sans-serif; background: #111318; color: #e2e2e9; display: grid; place-items: center; min-height: 100vh; padding: 24px; }
    .card { max-width: 680px; width: 100%; background: #191c22; border: 1px solid rgba(255,255,255,0.1); border-radius: 20px; padding: 28px 30px; box-shadow: 0 10px 30px rgba(0,0,0,0.35); }
    h1 { margin: 0 0 6px; font-size: 22px; }
    h2 { margin: 20px 0 6px; font-size: 15px; color: #8ab4f8; }
    p, li { color: #a8aebc; line-height: 1.6; font-size: 14px; }
    a { color: #8ab4f8; }
    .meta { font-size: 12px; color: #747b8a; margin-bottom: 18px; }
  </style>
</head>
<body>
  <main class="card">
    <h1>Fineprint — Privacy Policy &amp; Legal Disclaimer</h1>
    <div class="meta">Last updated: November 25, 2026 · <a href="/">← Back to Fineprint</a></div>

    <h2>1. Zero Contract Retention</h2>
    <p>Fineprint is a stateless Cloudflare Worker. Contract clauses and redlines sent to <code>https://fineprint.nimhq.com/mcp</code> are processed strictly in ephemeral memory to render the interactive UI and generate your exported file (<code>.docx</code>, <code>.pdf</code>, or <code>.md</code>). We do <strong>not</strong> log, store, persist, sell, or train AI models on any contract content.</p>

    <h2>2. No Accounts, Cookies, or Tracking</h2>
    <p>Fineprint requires no user login, sets no cookies, runs no third-party analytics, and maintains no database of users or documents. Standard transient edge network metadata (such as HTTP status codes) is handled by Cloudflare in accordance with Cloudflare's privacy policy.</p>

    <h2>3. Not Legal Advice</h2>
    <p>Fineprint is a document-formatting and drafting aid only. It is <strong>not</strong> a law firm, does not provide legal advice, and does not create an attorney-client relationship. Always have qualified legal counsel review any contract or redline before signing.</p>

    <h2>4. Open Source &amp; Contact</h2>
    <p>Fineprint is open source under the MIT License. You can audit the full source code or open an issue at <a href="https://github.com/nimeetshah0/fineprint" target="_blank" rel="noopener noreferrer">github.com/nimeetshah0/fineprint</a>.</p>
  </main>
</body>
</html>`;
}

export default {
  async fetch(request: Request): Promise<Response> {
    const url = new URL(request.url);

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }

    if (url.pathname === '/health') {
      return new Response(
        JSON.stringify({
          ok: true,
          service: 'fineprint',
          runtime: 'cloudflare-workers',
          transport: 'streamable-http',
          mcpEndpoint: '/mcp',
        }),
        { status: 200, headers: { 'Content-Type': 'application/json', ...CORS_HEADERS } }
      );
    }

    if (url.pathname === '/' && request.method === 'GET') {
      return new Response(landingHtml(url.origin), {
        status: 200,
        headers: { 'Content-Type': 'text/html; charset=utf-8' },
      });
    }

    if (
      (url.pathname === '/icon.svg' ||
        url.pathname === '/favicon.svg' ||
        url.pathname === '/favicon.ico') &&
      request.method === 'GET'
    ) {
      return new Response(ICON_SVG, {
        status: 200,
        headers: {
          'Content-Type': 'image/svg+xml',
          'Cache-Control': 'public, max-age=86400',
          ...CORS_HEADERS,
        },
      });
    }

    if (url.pathname === '/privacy' && request.method === 'GET') {
      return new Response(privacyHtml(), {
        status: 200,
        headers: { 'Content-Type': 'text/html; charset=utf-8' },
      });
    }

    if (url.pathname === '/mcp') {
      const mcpServer = createContractRedlinerServer({
        readAppHtml: () => String(APP_HTML),
        publicMode: true,
        listPlaybooks: () => BUILT_IN_PLAYBOOKS,
        getPlaybookById: (id) => BUILT_IN_PLAYBOOKS.find((p) => p.id === id),
      });

      const transport = new WebStandardStreamableHTTPServerTransport({
        sessionIdGenerator: undefined,
      });

      await mcpServer.connect(transport);
      const response = await transport.handleRequest(request);

      const headers = new Headers(response.headers);
      for (const [k, v] of Object.entries(CORS_HEADERS)) {
        headers.set(k, v);
      }

      return new Response(response.body, {
        status: response.status,
        statusText: response.statusText,
        headers,
      });
    }

    return new Response(JSON.stringify({ error: 'Not found. Use /mcp or /health.' }), {
      status: 404,
      headers: { 'Content-Type': 'application/json', ...CORS_HEADERS },
    });
  },
};
