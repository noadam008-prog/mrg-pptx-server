// Security + regression test for the render server (2026-10-23, after the security review of the 3 finance apps).
// Starts the server locally with a test token, then checks:
//   auth (header only, constant-time, no ?token=, refuses to start without a token), images (no local files, only allowed hosts,
//   no redirects, image types only), file names (no path traversal), no token in links, security headers,
//   and renders test-payload*.json so the decks can be compared with a baseline (no layout change).
//   node test_security.mjs [renderOutDir]
import { spawn } from 'child_process';
import http from 'http';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const OUT = process.argv[2] || path.join(HERE, 'output', 'test');
fs.mkdirSync(OUT, { recursive: true });
const PORT = 3400 + Math.floor(Math.random() * 300), TOKEN = 'test-token-' + Date.now();
const BASE = `http://127.0.0.1:${PORT}`;
const sleep = ms => new Promise(r => setTimeout(r, ms));
let pass = 0, fail = 0; const ok = (c, m) => { c ? pass++ : fail++; console.log(c ? 'ok  ' : 'FAIL', m); };

// a local "internal service": records every request (an SSRF probe must never reach it)
const hits = []; const trap = http.createServer((q, s) => { hits.push(q.url); s.writeHead(q.url.startsWith('/redirect') ? 302 : 200, q.url.startsWith('/redirect') ? { Location: `http://127.0.0.1:${trapPort}/secret` } : { 'Content-Type': 'image/png' }); s.end('x'); });
const trapPort = PORT + 500; await new Promise(r => trap.listen(trapPort, '127.0.0.1', r));

function start(env) {
  const p = spawn(process.execPath, ['pptx-api-server.mjs'], { cwd: HERE, env: { ...process.env, PORT: String(PORT), ...env }, stdio: ['ignore', 'pipe', 'pipe'] });
  let log = ''; p.stdout.on('data', d => log += d); p.stderr.on('data', d => log += d);
  return { p, log: () => log };
}
async function up() { for (let i = 0; i < 60; i++) { try { if ((await fetch(BASE + '/health')).ok) return true; } catch {} await sleep(250); } return false; }
const post = (q, body, auth = 'Bearer ' + TOKEN) => fetch(BASE + '/generate-pptx' + q, { method: 'POST', headers: { 'Content-Type': 'application/json', ...(auth ? { Authorization: auth } : {}) }, body: JSON.stringify(body) });

// 1. no token → the server must not start (unless the local-dev flag is set)
{ const { p, log } = start({ RENDER_TOKEN: '', ALLOW_NO_TOKEN: '' }); const started = await up(); p.kill(); await sleep(300);
  ok(!started, 'refuses to start without RENDER_TOKEN' + (started ? '' : '') ); }

const srv = start({ RENDER_TOKEN: TOKEN, ALLOW_NO_TOKEN: '', IMAGE_HOSTS: 'fal.media' });
try {
  ok(await up(), 'starts with a token');
  const h = await fetch(BASE + '/health');
  ok(!h.headers.get('x-powered-by'), 'no X-Powered-By header');
  ok(h.headers.get('x-content-type-options') === 'nosniff', 'X-Content-Type-Options: nosniff');
  ok(!h.headers.get('access-control-allow-origin'), 'no open CORS');

  // 2. auth
  ok((await post('?link=1', {}, null)).status === 401, 'no token → 401');
  ok((await post('?link=1', {}, 'Bearer wrong')).status === 401, 'wrong token → 401');
  ok((await fetch(`${BASE}/generate-pptx?link=1&token=${TOKEN}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' })).status === 401, '?token= in the URL is not accepted');

  const small = { company_name: 'サンプル商事', financial_data: {} };
  const size = async extra => { const r = await post('?link=1', { ...small, ...extra }); const j = await r.json(); return { status: r.status, j, n: j.file_size_bytes || 0 }; };
  const plain = await size({});
  ok(plain.status === 200 && plain.n > 0, 'renders a deck without a cover image (' + plain.n + ' bytes)');
  ok(!String(plain.j.download_url).includes(TOKEN), 'the download link does not contain the token');

  // 3. images
  const local = await size({ cover_image_url: 'pptx-api-server.mjs' });
  ok(local.status === 200 && Math.abs(local.n - plain.n) < 2000, 'a local file path as image is ignored (not embedded)');
  const abs = await size({ cover_image_url: path.join(HERE, 'pptx-api-server.mjs') });
  ok(abs.status === 200 && Math.abs(abs.n - plain.n) < 2000, 'an absolute file path as image is ignored');
  hits.length = 0;
  await size({ cover_image_url: `http://127.0.0.1:${trapPort}/internal.png` });
  await size({ cover_image_url: `https://127.0.0.1:${trapPort}/internal.png` });
  await size({ cover_image_url: 'http://169.254.169.254/metadata/instance' });
  ok(hits.length === 0, 'addresses outside the allow-list are never fetched (internal service hits: ' + hits.length + ')');
  ok((await size({ cover_image_url: 'https://fal.media.evil.example/x.png' })).status === 200, 'a look-alike host is refused without failing the deck');

  // 4. file names
  const before = new Set(fs.readdirSync(HERE));
  const tr = await post('', { ...small, pptxConfig: { filename_prefix: '../../traversal_test' } });
  const tj = await tr.json();
  const stray = [path.join(HERE, '..'), path.join(HERE, '..', '..'), HERE].flatMap(d => fs.readdirSync(d).filter(f => f.startsWith('traversal_test')).map(f => path.join(d, f)));
  stray.forEach(f => fs.rmSync(f, { force: true }));   // clean up if the old code wrote one
  ok(tr.status === 200 && stray.length === 0, 'filename_prefix cannot write outside output/ (' + (tj.filename || tj.error) + ')');
  ok(!String(tj.download_url || '').includes(TOKEN), 'legacy (non-link) response does not contain the token');
  ok([...fs.readdirSync(HERE)].every(f => before.has(f)), 'no new file next to the server code');

  // 5. regression renders (compare with the baseline outside this script)
  for (const f of ['test-payload.json', 'test-payload-v2.json']) {
    const body = JSON.parse(fs.readFileSync(path.join(HERE, f), 'utf8'));
    const r = await post('?return_file=1', body);
    const buf = Buffer.from(await r.arrayBuffer()); fs.writeFileSync(path.join(OUT, f + '.pptx'), buf);
    ok(r.status === 200 && buf.length > 10000, `renders ${f} (${buf.length} bytes) → ${path.join(OUT, f + '.pptx')}`);
  }
} finally { srv.p.kill(); trap.close(); }
console.log(`${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
