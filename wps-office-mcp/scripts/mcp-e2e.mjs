#!/usr/bin/env node
/**
 * 端到端测试：像 Claude Code 一样启动 dist/index.js（stdio），按顺序调用工具并打印结果。
 * 运行前要先停掉占用 58891 端口的其他 MCP 服务（同一时间只能有一个连 WPS）。
 *
 * 用法: node scripts/mcp-e2e.mjs <步骤文件.json>
 * 步骤文件: [{"tool":"wps_common_ping","args":{}}, {"listTools":true}, {"sleep":2000}, ...]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const steps = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));

const transport = new StdioClientTransport({
  command: process.execPath,
  args: [path.join(here, '..', 'dist', 'index.js')],
  stderr: 'ignore',
});
const client = new Client({ name: 'wps-e2e', version: '0.0.1' });
await client.connect(transport);

for (const s of steps) {
  if (s.sleep) { await new Promise(r => setTimeout(r, s.sleep)); continue; }
  if (s.listTools) {
    const { tools } = await client.listTools();
    const names = tools.map(t => t.name);
    console.log(`# 工具数: ${names.length}`);
    for (const n of s.expectMissing || []) console.log(`  ${names.includes(n) ? '✗ 仍在' : '✓ 已隐藏'} ${n}`);
    for (const n of s.expectPresent || []) console.log(`  ${names.includes(n) ? '✓ 存在' : '✗ 缺失'} ${n}`);
    continue;
  }
  const t0 = Date.now();
  let out;
  try {
    const r = await client.callTool({ name: s.tool, arguments: s.args || {} }, undefined, { timeout: 240000 });
    out = (r.isError ? '[isError] ' : '') + (r.content || []).map(c => c.text ?? JSON.stringify(c)).join('\n');
  } catch (e) {
    out = '[异常] ' + (e?.message || String(e));
  }
  const text = out.length > (s.max || 600) ? out.slice(0, s.max || 600) + ` …(共${out.length}字)` : out;
  console.log(`> ${s.note ? s.note + ' | ' : ''}${s.tool} ${JSON.stringify(s.args || {})}  (${Date.now() - t0}ms)`);
  console.log('  ' + text.replace(/\n/g, '\n  '));
}
await client.close();
process.exit(0);
