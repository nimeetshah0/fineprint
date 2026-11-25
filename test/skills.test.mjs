import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { z } from 'zod';
import { createContractRedlinerServer } from '../dist/server/mcpCore.js';
import { BUILT_IN_PLAYBOOKS } from '../dist/server/playbooks.js';

const server = createContractRedlinerServer({
  readAppHtml: () => '<html></html>',
  publicMode: true,
  listPlaybooks: () => BUILT_IN_PLAYBOOKS,
  getPlaybookById: (id) => BUILT_IN_PLAYBOOKS.find((p) => p.id === id),
});
const client = new Client({ name: 'skills-test', version: '1' });
const [clientT, serverT] = InMemoryTransport.createLinkedPair();
await Promise.all([server.connect(serverT), client.connect(clientT)]);

const ext = client.getServerCapabilities()?.extensions?.['io.modelcontextprotocol/skills'];
assert.deepEqual(ext, { directoryRead: false }, 'skills extension declared under capabilities.extensions');

const { skills } = await client.request({ method: 'skills/list', params: {} }, z.any());
assert.equal(skills.length, 1);
const [skill] = skills;
assert.equal(skill.uri, 'skill://fineprint/SKILL.md');

const { skill: got } = await client.request({ method: 'skills/get', params: { uri: skill.uri } }, z.any());
assert.deepEqual(got, skill);

for (const res of skill.resources) {
  const { contents } = await client.readResource({ uri: res.uri });
  const bytes = Buffer.from(contents[0].text, 'utf-8');
  assert.equal(bytes.byteLength, res.size, `${res.uri} size`);
  assert.equal(`sha256:${createHash('sha256').update(bytes).digest('hex')}`, res.digest, `${res.uri} digest`);
}

const { contents } = await client.readResource({ uri: skill.uri });
const fm = contents[0].text.match(/^---\nname: (.+)\ndescription: >-\n {2}(.+)\n---/);
assert.ok(fm, 'SKILL.md has frontmatter');
assert.equal(fm[1], skill.frontmatter.name);
assert.equal(fm[2], skill.frontmatter.description);
assert.equal(fm[1], skill.uri.split('/')[2], 'name matches skill directory');

await assert.rejects(
  client.request({ method: 'skills/get', params: { uri: 'skill://nope/SKILL.md' } }, z.any()),
  (err) => err.code === -32602
);

await client.close();
console.log('Skills over MCP: ALL CHECKS PASSED');
