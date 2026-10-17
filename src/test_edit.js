const fs = require('fs'), JSZip = require('jszip'), { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
require('/home/claude/mgr/core.js'); require('/home/claude/mgr/edit.js');
const C = globalThis.LeaveCore, E = globalThis.LeaveEdit;
const parseXML = s => new DOMParser().parseFromString(s, 'text/xml'), ser = d => new XMLSerializer().serializeToString(d);
fs.mkdirSync('edit_out', { recursive: true });
async function run(file, plan) {
  const src = '/home/claude/xl/files/' + file, cfg = C.deptForFile(file);
  const zip = await JSZip.loadAsync(fs.readFileSync(src));
  const book = await C.openWorkbook(JSZip, fs.readFileSync(src), parseXML);
  const before = await C.parseDept(cfg, book, [2026, 10]);
  const edits = [];
  for (const p of plan) {
    const R = before.roster[p.month]; const person = R.people.find(x => x.name.toUpperCase().includes(p.who));
    for (const d of p.days) {
      const e = { sheetPath: R.path, row: person.row, col: R.dayCol + d - 1, action: p.action };
      if (p.action === 'clear') { const other = R.people.find(o => o.group === person.group && o !== person && o.days[d - 1].code !== 'R' && !o.days[d - 1].text && !'DvYBO'.includes(o.days[d-1].code)); if (other) e.clearFillFrom = { row: other.row, col: e.col }; }
      edits.push(e);
    }
    console.log(`  ${p.action} ${person.name} ${p.month} giorni ${p.days.join(',')} (riga ${person.row}) prima: ${p.days.map(d => person.days[d - 1].code + (person.days[d - 1].text || '')).join(' ')}`);
  }
  const rep = await E.applyEdits(zip, parseXML, ser, edits, { author: 'Antonio Barone', when: new Date(2026, 9, 17, 23, 30) });
  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync('edit_out/' + file, out);
  const after = await C.parseDept(cfg, await C.openWorkbook(JSZip, out, parseXML), [2026, 10]);
  for (const p of plan) { const g = Object.keys(before.months[p.month].groups); for (const k of g) {
    const b = before.months[p.month].groups[k].booked, a = after.months[p.month].groups[k].booked;
    const ch = p.days.filter(d => b[d-1] !== a[d-1]).map(d => `${d}:${b[d-1]}→${a[d-1]}`); if (ch.length) console.log(`  conteggi ${k}: ${ch.join(' ')}`); } }
  console.log(`${file.slice(0, 20)}: celle ${rep.cells}, commenti nuovi ${rep.comments}, dimensione ${fs.statSync(src).size} → ${out.length}`);
}
(async () => {
  await run('HOST Leave Booking 2026&2027.xlsx', [
    { month: '2026-11', who: 'ISAK', days: [9, 10, 11], action: 'book' },
    { month: '2026-11', who: 'ELLA', days: [24], action: 'book' },
    { month: '2026-10', who: 'FRANCESCO', days: [1, 2], action: 'clear' },
    { month: '2026-10', who: 'CATALINA', days: [28], action: 'cancel' },
  ]);
  await run('SEC&DRIVER Leave Booking 2026&2027.xlsx', [
    { month: '2027-05', who: 'ZIKIC', days: [3, 4, 5, 6, 7], action: 'book' },
    { month: '2027-05', who: 'GRECH', days: [10], action: 'book' },
  ]);
})().catch(e => console.log('ERRORE', e));
