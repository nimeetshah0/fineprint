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

## Contributing Playbooks (via GitHub PR)

All built-in playbooks live in [`src/playbooks.json`](src/playbooks.json). To add a new playbook (e.g., Commercial Lease, Healthcare BAA, Consulting Agreement) or improve existing rules:

1. Edit [`src/playbooks.json`](https://github.com/nimeetshah0/fineprint/edit/main/src/playbooks.json) directly on GitHub or in a fork.
2. Add an entry matching this shape:
   ```json
   {
     "id": "commercial-lease",
     "name": "Commercial Office Lease",
     "audience": "Founders & Ops",
     "description": "Flags personal guarantees, uncapped CAM charges, and renewal traps.",
     "builtIn": true,
     "rules": [
       {
         "id": "lease-guarantee",
         "topic": "Personal Guarantee",
         "severity": "dealbreaker",
         "position": "Remove personal guarantees; limit security to corporate deposit or letter of credit.",
         "fallback": "Good-guy guarantee capped at 3 months rent."
       }
     ]
   }
   ```
3. Open a [Pull Request](https://github.com/nimeetshah0/fineprint/pulls) — once merged, it goes live automatically on `https://fineprint.nimhq.com/mcp`.

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
