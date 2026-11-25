#!/usr/bin/env node
import crypto from 'node:crypto';
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { parseDocxFile } from './docxEngine.js';
import { createContractRedlinerServer, type InMemorySession } from './mcpCore.js';
import { parsePdfFile } from './pdfEngine.js';
import { BUILT_IN_PLAYBOOKS, PlaybookStore } from './playbooks.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function getArgValue(flag: string): string | undefined {
  const eqPrefix = `${flag}=`;
  for (let i = 0; i < process.argv.length; i++) {
    const arg = process.argv[i];
    if (arg.startsWith(eqPrefix)) return arg.slice(eqPrefix.length);
    if (arg === flag && process.argv[i + 1]) return process.argv[i + 1];
  }
  return undefined;
}

const dataDir = getArgValue('--data_dir') || process.env.CONTRACT_REDLINER_DATA_DIR;
const playbookStore = new PlaybookStore(dataDir);
const sessions = new Map<string, InMemorySession>();

function resolveAppHtmlPath(): string {
  if (process.env.CONTRACT_REDLINER_APP_HTML) {
    return process.env.CONTRACT_REDLINER_APP_HTML;
  }
  return path.resolve(__dirname, '..', 'ui', 'index.html');
}

function readAppHtml(): string {
  const htmlPath = resolveAppHtmlPath();
  if (!fs.existsSync(htmlPath)) {
    return `<!doctype html><html><body><h3>Contract Redliner UI not built</h3><p>${htmlPath}</p></body></html>`;
  }
  return fs.readFileSync(htmlPath, 'utf-8');
}

async function readLocalContract(filePath: string) {
  const expanded = filePath.startsWith('~')
    ? path.join(os.homedir(), filePath.slice(1))
    : filePath;
  const resolved = path.resolve(expanded);
  const ext = path.extname(resolved).toLowerCase();
  const sessionId = `sess_${crypto.randomBytes(6).toString('hex')}`;

  if (ext === '.pdf') {
    const parsed = await parsePdfFile(resolved);
    sessions.set(sessionId, {
      sessionId,
      title: parsed.fileName.replace(/\.pdf$/i, ''),
      counterparty: '',
      playbookId: 'startup-vendor-msa',
      sourceType: 'pdf',
      sourceFileName: parsed.fileName,
      sourceFilePath: resolved,
      clauses: [],
    });

    const numbered = parsed.paragraphs
      .map((p) => `[Page ${p.page} · P${p.index}] ${p.text}`)
      .join('\n\n');

    return {
      content: [
        {
          type: 'text' as const,
          text: `Loaded PDF \`${parsed.fileName}\` in memory (**sessionId:** \`${sessionId}\`, ${parsed.pageCount} page(s), ${parsed.paragraphs.length} sections). Pass this \`sessionId\` to \`open_redliner\` — exported redlines will generate a clean Word (\`.docx\`) file with native Track Changes and Margin Comments.\n\n${numbered}`,
        },
      ],
      structuredContent: {
        sessionId,
        fileName: parsed.fileName,
        sourceType: 'pdf',
        pageCount: parsed.pageCount,
        paragraphCount: parsed.paragraphs.length,
      },
    };
  }

  const parsed = await parseDocxFile(resolved);
  sessions.set(sessionId, {
    sessionId,
    title: parsed.fileName.replace(/\.docx$/i, ''),
    counterparty: '',
    playbookId: 'startup-vendor-msa',
    sourceType: 'docx',
    sourceFileName: parsed.fileName,
    sourceFilePath: resolved,
    sourceDocx: parsed,
    clauses: [],
  });

  const numbered = parsed.paragraphs.map((p) => `[P${p.index}] ${p.text}`).join('\n\n');

  return {
    content: [
      {
        type: 'text' as const,
        text: `Loaded \`${parsed.fileName}\` in memory (**sessionId:** \`${sessionId}\`, ${parsed.paragraphs.length} paragraphs). Pass this \`sessionId\` and matching \`paragraphIndex\` values to \`open_redliner\` so exported Track Changes preserve original Word formatting.\n\n${numbered}`,
      },
    ],
    structuredContent: {
      sessionId,
      fileName: parsed.fileName,
      sourceType: 'docx',
      paragraphCount: parsed.paragraphs.length,
    },
  };
}

function buildServer(publicMode: boolean) {
  return createContractRedlinerServer({
    readAppHtml,
    publicMode,
    listPlaybooks: () => (publicMode ? BUILT_IN_PLAYBOOKS : playbookStore.list()),
    getPlaybookById: (id) =>
      publicMode
        ? BUILT_IN_PLAYBOOKS.find((p) => p.id === id)
        : playbookStore.getById(id),
    saveCustomPlaybook: publicMode ? undefined : (pb) => playbookStore.upsertCustom(pb),
    sessions: publicMode ? undefined : sessions,
    readLocalContract: publicMode ? undefined : readLocalContract,
    rehydrateLocalDocx: publicMode
      ? undefined
      : async (filePath) => {
          if (!filePath.toLowerCase().endsWith('.docx') || !fs.existsSync(filePath)) {
            return undefined;
          }
          try {
            return await parseDocxFile(filePath);
          } catch {
            return undefined;
          }
        },
    saveToDownloads: publicMode
      ? undefined
      : (fileName, buffer) => {
          try {
            const dlDir = path.join(os.homedir(), 'Downloads');
            if (fs.existsSync(dlDir)) {
              const savedPath = path.join(dlDir, fileName);
              fs.writeFileSync(savedPath, buffer);
              return savedPath;
            }
          } catch {
            // fallback to browser download only
          }
          return undefined;
        },
  });
}

function readJsonBody(req: http.IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    req.on('data', (c: Buffer) => chunks.push(c));
    req.on('end', () => {
      const raw = Buffer.concat(chunks).toString('utf-8').trim();
      if (!raw) return resolve(undefined);
      try {
        resolve(JSON.parse(raw));
      } catch (err) {
        reject(err);
      }
    });
    req.on('error', reject);
  });
}

function setCorsHeaders(res: http.ServerResponse) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  res.setHeader(
    'Access-Control-Allow-Headers',
    'Content-Type, Accept, Mcp-Session-Id, Authorization'
  );
  res.setHeader('Access-Control-Expose-Headers', 'Mcp-Session-Id');
}

async function startHttpServer(port: number, host: string) {
  const httpServer = http.createServer(async (req, res) => {
    setCorsHeaders(res);
    const urlPath = (req.url || '/').split('?')[0];

    if (req.method === 'OPTIONS') {
      res.writeHead(204);
      res.end();
      return;
    }

    if (urlPath === '/health' || urlPath === '/') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(
        JSON.stringify({
          ok: true,
          service: 'contract-redliner',
          transport: 'streamable-http',
          publicMode: true,
          mcpEndpoint: '/mcp',
        })
      );
      return;
    }

    if (urlPath === '/mcp') {
      try {
        const body = req.method === 'POST' ? await readJsonBody(req) : undefined;
        const mcpServer = buildServer(true);
        const transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: undefined,
        });

        res.on('close', () => {
          void transport.close();
          void mcpServer.close();
        });

        await mcpServer.connect(transport);
        await transport.handleRequest(req, res, body);
      } catch (err) {
        if (!res.headersSent) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(
            JSON.stringify({
              jsonrpc: '2.0',
              error: {
                code: -32603,
                message: err instanceof Error ? err.message : 'Internal server error',
              },
              id: null,
            })
          );
        }
      }
      return;
    }

    res.writeHead(404, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Not found. Use /mcp or /health.' }));
  });

  httpServer.listen(port, host, () => {
    console.error(
      `[contract-redliner] Public Streamable HTTP MCP server listening on http://${host}:${port}/mcp`
    );
  });
}

async function main() {
  const isHttp =
    process.argv.includes('--http') ||
    process.env.MCP_TRANSPORT === 'http' ||
    Boolean(process.env.PORT);

  if (isHttp) {
    const port = Number(getArgValue('--port') || process.env.PORT || 8787);
    const host = getArgValue('--host') || process.env.HOST || '0.0.0.0';
    await startHttpServer(port, host);
    return;
  }

  const server = buildServer(false);
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  console.error('contract-redliner server error:', err);
  process.exit(1);
});
