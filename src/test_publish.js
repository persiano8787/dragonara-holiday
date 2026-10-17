const { chromium } = require('playwright'); const http = require('http'), fs = require('fs');
fs.copyFileSync('/home/claude/mgr/manager.html', '/tmp/claude-0/manager_test.html');
const srv = http.createServer((q, r) => { r.writeHead(200, { 'content-type': 'text/html' }); r.end(fs.readFileSync('/tmp/claude-0/manager_test.html')); }).listen(8766);
const HOST = '/home/claude/xl/files/HOST Leave Booking 2026&2027.xlsx';
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell' });
  const pg = await b.newPage({ viewport: { width: 1366, height: 860 } }); const errs = []; pg.on('pageerror', e => errs.push(e.message));
  await pg.route(/cdnjs\.cloudflare\.com.*jszip/, rt => rt.fulfill({ body: fs.readFileSync('node_modules/jszip/dist/jszip.min.js'), contentType: 'application/javascript' }));
  await pg.route(/fonts\.(googleapis|gstatic)/, rt => rt.abort());
  const puts = []; let tokenSeen = null;
  await pg.route(/api\.github\.com/, async rt => { const q = rt.request(); tokenSeen = q.headers()['authorization'];
    if (q.method() === 'PUT') { puts.push({ url: q.url(), body: JSON.parse(q.postData()) }); return rt.fulfill({ status: 200, json: { content: {} } }); }
    if (/\/contents\//.test(q.url())) return rt.fulfill({ status: 200, json: { sha: 'abc123' } });
    return rt.fulfill({ status: 200, json: { full_name: 'persiano8787/dragonara-holiday', permissions: { push: true } } }); });
  await pg.addInitScript(() => { window.showDirectoryPicker = async () => (await navigator.storage.getDirectory()).getDirectoryHandle('ferie', { create: true }); });
  await pg.goto('http://localhost:8766/');
  // copia il file Host nella cartella del browser
  await pg.evaluate(async b64 => { const dir = await (await navigator.storage.getDirectory()).getDirectoryHandle('ferie', { create: true });
    const h = await dir.getFileHandle('HOST Leave Booking 2026&2027.xlsx', { create: true }); const w = await h.createWritable();
    await w.write(Uint8Array.from(atob(b64), c => c.charCodeAt(0))); await w.close(); }, fs.readFileSync(HOST).toString('base64'));
  console.log('1 stato iniziale:', await pg.textContent('#pubSettings'), '| pulsante:', await pg.textContent('#save'));
  await pg.click('#pubSettings'); await pg.fill('#mTok', 'github_pat_TEST'); await pg.click('#mYes'); await pg.waitForTimeout(1300);
  console.log('2 dopo token:', await pg.textContent('#pubSettings'), '| pulsante:', await pg.textContent('#save'), '| header inviato:', tokenSeen);
  await pg.click('#pickDir'); await pg.waitForSelector('.tab'); await pg.click('#next');
  await pg.fill('#author', 'Antonio Barone');
  const row = pg.locator('#grid tbody tr:has(td.name:text-is("ISAK")) td.c');
  // trova un giorno modificabile con almeno 1 posto libero
  const free = await pg.$$eval('tr.free td.c', tds => tds.map(t => t.textContent));
  const ed = await row.evaluateAll(tds => tds.map(t => !!t.dataset.k));
  const day = free.findIndex((v, i) => +v >= 1 && ed[i]);
  await row.nth(day).click(); await pg.click('#actBook'); await pg.waitForTimeout(600);
  await pg.click('#save'); await pg.waitForSelector('#saveLog:not([hidden])'); await pg.waitForTimeout(300);
  console.log('3 log:', (await pg.textContent('#logList')).trim());
  const p = puts[0]; const json = Buffer.from(p.body.content, 'base64').toString('utf8'); const data = JSON.parse(json);
  console.log('4 PUT su:', p.url.replace(/.*repos\//, ''), '| messaggio:', p.body.message, '| sha:', p.body.sha, '| branch:', p.body.branch);
  console.log('   contenuto: reparto', data.dept, '| mesi', Object.keys(data.months).length, '| gruppi', data.groups.map(g => g.id).join(','), '| updated', data.updated.slice(0, 16));
  const names = ['ISAK', 'ELIZABETH', 'ALEXIA', 'SUSMA', 'CATALINA', 'Antonio', 'Barone'];
  console.log('   nomi nel contenuto pubblicato:', names.filter(n => json.toUpperCase().includes(n.toUpperCase())).join(',') || 'nessuno', '| byte', json.length);
  const files = await pg.evaluate(async () => { const d = await (await navigator.storage.getDirectory()).getDirectoryHandle('ferie'); const out = []; for await (const [n, h] of d.entries()) { if (h.kind === 'directory') for await (const [m] of h.entries()) out.push(n + '/' + m); else out.push(n); } return out; });
  console.log('5 cartella dopo il salvataggio:', files.join(' | '));
  // Publish senza modifiche
  await pg.click('#publishAll'); await pg.click('#mYes'); await pg.waitForTimeout(800);
  console.log('6 Publish:', (await pg.textContent('#logList')).trim(), '| PUT totali:', puts.length);
  await pg.screenshot({ path: 'shot_pub.png' });
  console.log('errori JS:', errs.length ? errs : 'nessuno'); await b.close(); srv.close();
})().catch(e => { console.log('ERRORE', e.message); process.exit(1); });
