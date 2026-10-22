#!/usr/bin/env node
/**
 * Mac 加载项直连测试工具（开发用，不经过 MCP）。
 * 在 127.0.0.1:58891 扮演 MCP 服务端，把命令直接发给 WPS 加载项，打印原始返回。
 * 运行前要先停掉占用 58891 的 MCP 服务。
 *
 * 用法: node scripts/mac-harness.mjs <命令文件.json> [--wait 秒]
 * 命令文件: [{"app":"et|wps|wpp|any","action":"getAppInfo","params":{}}, ...]
 *   app 省略或为 any 时，发给第一个来取的加载项（旧版加载项不带 app 参数，只能用 any）。
 */
import http from 'node:http';
import fs from 'node:fs';

const file = process.argv[2];
const waitIdx = process.argv.indexOf('--wait');
const startWait = waitIdx > 0 ? Number(process.argv[waitIdx + 1]) : 15;
const cmds = JSON.parse(fs.readFileSync(file, 'utf8'));

const lastPoll = {};          // app -> {t, doc}
let pending = null;           // {app, action, params, requestId, resolve, sent}

const server = http.createServer((req, res) => {
  const origin = req.headers.origin;
  if (origin && /^https?:\/\//i.test(origin)) { res.writeHead(403); res.end(); return; }
  res.setHeader('Access-Control-Allow-Origin', origin || '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Content-Type', 'application/json');
  if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }
  const url = new URL(req.url, 'http://x');
  if (url.pathname === '/poll') {
    const app = url.searchParams.get('app') || 'any';
    const inst = url.searchParams.get('inst') || '-';
    lastPoll[app + '/' + inst] = { t: Date.now(), doc: url.searchParams.get('doc') || '' };
    if (pending && !pending.sent && (pending.app === 'any' || pending.app === app)) {
      pending.sent = app + '/' + inst;
      res.end(JSON.stringify({ command: { action: pending.action, params: pending.params, requestId: pending.requestId, app: pending.app } }));
    } else {
      res.end('{}');
    }
    return;
  }
  if (url.pathname === '/result' && req.method === 'POST') {
    let body = '';
    req.on('data', c => (body += c));
    req.on('end', () => {
      const data = JSON.parse(body);
      if (pending && data.requestId === pending.requestId) {
        const p = pending; pending = null; p.resolve({ by: p.sent, result: data.result });
      } else {
        console.log('  [迟到/重复结果]', data.requestId, JSON.stringify(data.result).slice(0, 200));
      }
      res.end('{"ok":true}');
    });
    return;
  }
  res.writeHead(404); res.end('{}');
});

function send(c, timeoutMs = 60000) {
  return new Promise((resolve) => {
    const requestId = 'h-' + Date.now() + '-' + Math.random().toString(36).slice(2, 7);
    pending = { app: c.app || 'any', action: c.action, params: c.params || {}, requestId, resolve, sent: null };
    setTimeout(() => { if (pending?.requestId === requestId) { pending = null; resolve({ timeout: true }); } }, c.timeout || timeoutMs);
  });
}

server.listen(58891, '127.0.0.1', async () => {
  const t0 = Date.now();
  while (Object.keys(lastPoll).length === 0 && Date.now() - t0 < startWait * 1000) await new Promise(r => setTimeout(r, 200));
  await new Promise(r => setTimeout(r, 1200));   // 让各组件都来一轮
  console.log('在线:', JSON.stringify(Object.fromEntries(Object.entries(lastPoll).map(([k, v]) => [k, { ago: Date.now() - v.t, doc: v.doc }]))));
  for (const c of cmds) {
    const r = await send(c);
    console.log(`> ${c.app || 'any'} ${c.action} ${JSON.stringify(c.params || {})}`);
    console.log('  ', JSON.stringify(r).slice(0, 1500));
    if (c.pause) await new Promise(r => setTimeout(r, c.pause));
  }
  await new Promise(r => setTimeout(r, 1500));   // 收迟到结果
  server.close(); process.exit(0);
});
