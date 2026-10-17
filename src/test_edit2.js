const fs = require('fs'), JSZip = require('jszip'), { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
require('/home/claude/mgr/core.js'); require('/home/claude/mgr/edit.js');
const C = globalThis.LeaveCore, E = globalThis.LeaveEdit;
const parseXML = s => new DOMParser().parseFromString(s, 'text/xml'), ser = d => new XMLSerializer().serializeToString(d);
fs.mkdirSync('edit_out', { recursive: true });
(async () => {
  for (const [file, month, who, days, author] of [
    ['F&B Gaming Leave Booking 2026&2027.xlsx', '2027-08', 'JENNY', [10, 11, 12], 'Test Manager'],
    ['HOST Leave Booking 2026&2027.xlsx', '2026-11', 'ISAK', [9, 10, 11], 'Antonio Barone'],
    ['SEC&DRIVER Leave Booking 2026&2027.xlsx', '2027-05', 'ZIKIC', [3, 4, 5, 6, 7], 'Antonio Barone']]) {
    const src = '/home/claude/xl/files/' + file, cfg = C.deptForFile(file);
    const book = await C.openWorkbook(JSZip, fs.readFileSync(src), parseXML);
    const before = await C.parseDept(cfg, book, [2026, 10]);
    const R = before.roster[month], person = R.people.find(x => x.name.toUpperCase().includes(who));
    const zip = await JSZip.loadAsync(fs.readFileSync(src));
    const rep = await E.applyEdits(zip, parseXML, ser, days.map(d => ({ sheetPath: R.path, row: person.row, col: R.dayCol + d - 1, action: 'book' })), { author, when: new Date(2026, 9, 17, 23, 45) });
    const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    fs.writeFileSync('edit_out/' + file, out);
    const names = Object.keys((await JSZip.loadAsync(out)).files).filter(n => n.endsWith('/'));
    console.log(file.slice(0, 18), month, person.name, '| celle', rep.cells, 'commenti', rep.comments, '| cartelle nello zip:', names.length);
  }
})().catch(e => console.log('ERRORE', e));
