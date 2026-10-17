const fs = require('fs'), path = require('path');
const JSZip = require('jszip'); const { DOMParser } = require('@xmldom/xmldom');
require('/home/claude/mgr/core.js'); const C = globalThis.LeaveCore;
const parseXML = s => new DOMParser().parseFromString(s, 'text/xml');
const py = JSON.parse(fs.readFileSync('/home/claude/scripts/availability_all.json'));
const pyName = { "Cashdesk": "Cashdesk", "F&B": "F&B Gaming", "Host & Reception": "Host", "Gaming & Poker": "LG", "Security & Driver": "Security & Driver" };
(async () => {
  const dir = '/home/claude/xl/files';
  for (const f of fs.readdirSync(dir)) {
    const cfg = C.deptForFile(f); if (!cfg) { console.log('ignorato:', f); continue; }
    const t0 = Date.now();
    const book = await C.openWorkbook(JSZip, fs.readFileSync(path.join(dir, f)), parseXML);
    const res = await C.parseDept(cfg, book, [2026, 10]);
    const ref = py[pyName[cfg.dept]]; let diff = 0, cmp = 0;
    for (const [mk, m] of Object.entries(res.months)) {
      const r = ref[mk]; if (!r) { diff++; continue; }
      if (JSON.stringify(m.blocked) !== JSON.stringify(r.blocked)) { diff++; console.log('  blocked', mk); }
      for (const [g, v] of Object.entries(m.groups)) {
        const rg = r.groups[g === 'Host & Reception' ? 'Host' : g];
        cmp++;
        if (!rg || rg.limit !== v.limit || JSON.stringify(rg.booked) !== JSON.stringify(v.booked)) { diff++; console.log('  diff', mk, g, rg && rg.booked.join(''), v.booked.join('')); }
      }
    }
    const kinds = res.issues.reduce((a, i) => (a[i.kind] = (a[i.kind] || 0) + 1, a), {});
    console.log(`${cfg.dept.padEnd(18)} mesi ${Object.keys(res.months).length}  gruppi-mese confrontati ${cmp}  differenze ${diff}  segnalazioni ${JSON.stringify(kinds)}  ${Date.now() - t0}ms`);
  }
})().catch(e => console.log('ERRORE', e));
