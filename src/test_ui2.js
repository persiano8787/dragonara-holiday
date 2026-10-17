const fs = require('fs'), path = require('path'); const { JSDOM } = require('jsdom');
const html = fs.readFileSync('/home/claude/mgr/manager.html', 'utf8').replace(/<script src="https:\/\/cdnjs[^>]+><\/script>/, '').replace(/<link[^>]+fonts[^>]+>/g, '');
const dom = new JSDOM(html, { runScripts: 'dangerously', pretendToBeVisual: true, url: 'https://example.org/' });
const w = dom.window, d = w.document; const errs = []; w.addEventListener('error', e => errs.push(e.message));
w.JSZip = require('jszip'); w.Uint8Array = Uint8Array;
// ---- fake File System Access handles over a real folder
const FOLDER = path.resolve('fakefolder');
function fileHandle(dirPath, name){ const p = path.join(dirPath, name); return { kind: 'file', name,
  async getFile(){ const b = fs.readFileSync(p); const st = fs.statSync(p); return { name, lastModified: Math.floor(st.mtimeMs), arrayBuffer: async () => b }; },
  async createWritable(){ if (fs.existsSync(path.join(dirPath, '~$' + name))) throw new Error('locked'); let chunks = []; return { write: async x => chunks.push(Buffer.from(x instanceof ArrayBuffer ? new Uint8Array(x) : x)), close: async () => fs.writeFileSync(p, Buffer.concat(chunks)) }; } }; }
function dirHandle(dirPath){ return { kind: 'directory', name: path.basename(dirPath),
  async queryPermission(){ return 'granted'; }, async requestPermission(){ return 'granted'; },
  async getFileHandle(n, o = {}){ const p = path.join(dirPath, n); if (!fs.existsSync(p)) { if (!o.create) { const e = new Error('NotFound'); e.name = 'NotFoundError'; throw e; } fs.writeFileSync(p, ''); } return fileHandle(dirPath, n); },
  async getDirectoryHandle(n, o = {}){ const p = path.join(dirPath, n); if (!fs.existsSync(p)) fs.mkdirSync(p); return dirHandle(p); },
  async removeEntry(n){ fs.unlinkSync(path.join(dirPath, n)); },
  async *entries(){ for (const n of fs.readdirSync(dirPath)) yield [n, fs.statSync(path.join(dirPath, n)).isDirectory() ? dirHandle(path.join(dirPath, n)) : fileHandle(dirPath, n)]; } }; }
w.showDirectoryPicker = async () => dirHandle(FOLDER);
const sleep = ms => new Promise(r => setTimeout(r, ms));
const keep = setInterval(() => {}, 1000);
const tabByName = n => [...d.querySelectorAll('.tab')].find(t => t.textContent.startsWith(n));
const freeRow = () => [...d.querySelectorAll('tr.free')].map(r => [...r.querySelectorAll('td.c')].map(c => c.textContent).join(''));
const mouse = (el, type, opts = {}) => el.dispatchEvent(new w.MouseEvent(type, { bubbles: true, ...opts }));
async function waitFor(fn, ms = 20000){ const t = Date.now(); while (!fn()) { if (Date.now() - t > ms) throw new Error('timeout'); await sleep(50); } }
(async () => {
  await sleep(100);
  d.getElementById('pickDir').click(); await waitFor(() => d.querySelectorAll('.tab').length === 5);
  console.log('1 caricati:', d.getElementById('status').textContent, '| file:', [...d.querySelectorAll('#files li')].map(l => l.textContent.split(' · ')[1]).join(', '));
  tabByName('Host').click(); d.getElementById('next').click(); // November 2026
  console.log('2 mese:', d.getElementById('month').textContent, '| liberi prima:', freeRow()[0].slice(0, 12));
  // seleziona ISAK 9..11 con trascinamento
  const row = [...d.querySelectorAll('#grid tbody tr')].find(r => r.querySelector('td.name').textContent.startsWith('ISAK'));
  const cells = [...row.querySelectorAll('td.c')];
  mouse(cells[8], 'mousedown'); mouse(cells[10], 'mouseover'); mouse(d, 'mouseup');
  console.log('3 selezione:', d.getElementById('selInfo').textContent, '| Book attivo senza nome?', !d.getElementById('actBook').disabled);
  d.getElementById('author').value = 'Antonio Barone'; d.getElementById('author').dispatchEvent(new w.Event('input'));
  d.getElementById('actBook').click(); await waitFor(() => !d.getElementById('modal').hidden);
  console.log('4 avviso limite:', d.getElementById('modalBox').textContent.replace(/\s+/g, ' ').slice(0, 120));
  d.getElementById('mYes').click(); await waitFor(() => d.getElementById('pending').textContent);
  console.log('5 liberi dopo:', freeRow()[0].slice(0, 12), '|', d.getElementById('pending').textContent, '| celle segnate:', d.querySelectorAll('td.changed').length);
  d.getElementById('actUndo').click(); await waitFor(() => !d.getElementById('pending').textContent);
  console.log('6 dopo Undo liberi:', freeRow()[0].slice(0, 12), '| in sospeso:', JSON.stringify(d.getElementById('pending').textContent));
  // prenota 1 giorno libero senza superare il limite, poi Cancel e Clear su altre celle
  const row2 = [...d.querySelectorAll('#grid tbody tr')].find(r => r.querySelector('td.name').textContent.startsWith('ISAK'));
  const free = freeRow()[0]; const day = [...free].findIndex((v, i) => +v >= 1 && row2.querySelectorAll('td.c')[i].dataset.k) + 1;
  mouse(row2.querySelectorAll('td.c')[day - 1], 'mousedown'); mouse(d, 'mouseup');
  d.getElementById('actBook').click(); await waitFor(() => d.getElementById('pending').textContent);
  console.log('7 prenotato giorno', day, 'senza avviso | liberi:', freeRow()[0].slice(0, 12), '| modale aperta?', !d.getElementById('modal').hidden);
  // stale check: tocca il file su disco
  const hostFile = path.join(FOLDER, 'HOST Leave Booking 2026&2027.xlsx');
  const t0 = fs.statSync(hostFile).mtime; fs.utimesSync(hostFile, new Date(), new Date(Date.now() + 5000));
  d.getElementById('save').click(); await waitFor(() => !d.getElementById('saveLog').hidden);
  console.log('8 salvataggio con file cambiato:', d.getElementById('logList').textContent);
  fs.utimesSync(hostFile, t0, t0);
  // lock check
  fs.writeFileSync(path.join(FOLDER, '~$HOST Leave Booking 2026&2027.xlsx'), 'x');
  d.getElementById('saveLog').hidden = true; d.getElementById('save').click(); await waitFor(() => !d.getElementById('saveLog').hidden);
  console.log('9 salvataggio con file aperto in Excel:', d.getElementById('logList').textContent);
  fs.unlinkSync(path.join(FOLDER, '~$HOST Leave Booking 2026&2027.xlsx'));
  d.getElementById('saveLog').hidden = true; d.getElementById('save').click(); await waitFor(() => !d.getElementById('saveLog').hidden);
  console.log('10 salvataggio normale:', d.getElementById('logList').textContent, '| backup:', fs.readdirSync(path.join(FOLDER, '_backup')));
  console.log('   in sospeso dopo:', JSON.stringify(d.getElementById('pending').textContent), '| Save attivo?', !d.getElementById('save').disabled);
  console.log('errori JS:', errs.length ? errs : 'nessuno'); clearInterval(keep); process.exit(0);
})().catch(e => { console.log('ERRORE', e, errs); process.exit(1); });
