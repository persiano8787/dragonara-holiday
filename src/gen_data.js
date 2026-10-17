const fs=require('fs'),JSZip=require('jszip'),{DOMParser}=require('@xmldom/xmldom');require('/home/claude/mgr/core.js');const C=globalThis.LeaveCore;
(async()=>{const dir='/home/claude/xl/files', now=new Date().toISOString();
for(const f of fs.readdirSync(dir)){const cfg=C.deptForFile(f);if(!cfg)continue;
 const b=await C.openWorkbook(JSZip,fs.readFileSync(dir+'/'+f),s=>new DOMParser().parseFromString(s,'text/xml'));
 const r=await C.parseDept(cfg,b,[2026,10]);
 fs.writeFileSync(`/home/claude/dragonara-holiday/data/${cfg.id}.json`,JSON.stringify(C.publicData(cfg,r.months,now)));
 console.log(cfg.id,Object.keys(r.months).length,'mesi');}})();
