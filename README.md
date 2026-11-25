# Fineprint — Contract Redliner MCP App

Catch hidden contract traps before you sign. A free, open-source **MCP App (SEP-1865)** that turns **Claude**, **ChatGPT**, and **Google Antigravity** into an interactive Material 3 contract redlining workbench—with **native Microsoft Word (`.docx`) Track Changes & Margin Comments** export.

**Live Hosted MCP Endpoint:** `https://fineprint.nimhq.com/mcp`

- **Zero Contract Retention:** Contracts are processed strictly in memory and never stored on disk.
- **Zero Extra LLM Cost:** Uses your existing Claude, ChatGPT, or Antigravity subscription to analyze clauses against built-in playbooks.
- **Bulletproof OOXML `.docx` Fidelity:** Preserves run fonts, tables, headers/footers, bookmarks, live Word `REF` cross-reference fields, existing counterparty Track Changes, and existing margin comments.
- **4 Export Formats:**
  1. **Word — Track Changes (`.docx`)** (`<w:del>` / `<w:ins>` + `word/comments.xml` bubbles)
  2. **Word — Clean Final (`.docx`)** (accepted edits applied cleanly, ready to sign)
  3. **PDF — Redline Report (`.pdf`)** (shareable executive summary)
  4. **Markdown — Memo (`.md`)** (for Notion, Slack, or email)

---

## Connect via Hosted URL (Claude.ai, ChatGPT, Mobile)

Add your deployed Streamable HTTP endpoint (`https://<your-subdomain>/mcp`):

- **Claude.ai / Claude Mobile:** *Settings → Connectors → Add Custom Connector* → paste `https://<your-subdomain>/mcp`
- **ChatGPT:** *Settings → Connectors / Apps → Add MCP Server* → paste `https://<your-subdomain>/mcp`

## Run Locally (Full `.docx` & `.pdf` File Round-Tripping)

Running locally via `npx` unlocks `read_contract` so you can point your agent at local `.docx` or `.pdf` files and export formatting-preserving Track Changes back to `~/Downloads`.

### Claude Desktop (`claude_desktop_config.json`)
```json
{
  "mcpServers": {
    "contract-redliner": {
      "command": "npx",
      "args": ["-y", "contract-redliner-mcp"]
    }
  }
}
```

### Google Antigravity (`~/.gemini/config/plugins/contract-redliner/mcp_config.json`)
```json
{
  "mcpServers": {
    "redliner": {
      "command": "npx",
      "args": ["-y", "contract-redliner-mcp", "--data_dir=${PLUGIN_DATA}"]
    }
  }
}
```

---

## Deploy Your Own to Cloudflare Workers (Free Tier)

```bash
npm install
npm run build
# Optional: set your custom subdomain in wrangler.toml under `routes`
npm run deploy:cf
```

## Development & Fidelity Tests

```bash
npm run build
npm run test:docx   # Runs 17 OOXML fidelity checks (including a 40+ page Enterprise MSA)
```

## License

MIT
