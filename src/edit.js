/* Holiday Availability — surgical .xlsx editor.
   Changes only: the fill of the touched cells (a new/reused cell style that keeps
   borders, fonts and alignment), the text of cleared cells, and new threaded comments
   (modern Excel comments, with their legacy fallback note and VML shape).
   Every other part of the package is written back byte-for-byte as it was read.     */
(function (root) {
  "use strict";
  const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
  const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
  const TC_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
  const T = {
    comments: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    vml: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
    threaded: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
    person: "http://schemas.microsoft.com/office/2017/10/relationships/person",
  };
  const CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
  const CT_THREADED = "application/vnd.ms-excel.threadedcomments+xml";
  const CT_PERSON = "application/vnd.ms-excel.person+xml";
  const CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing";
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
  const FILL = { book: "FF92D050", cancel: "FFFF0000" };
  const LEGACY_TEXT = "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    ";

  const guid = () => "{" + "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, c => {
    const r = Math.random() * 16 | 0; return (c === "x" ? r : (r & 3 | 8)).toString(16);
  }).toUpperCase() + "}";
  const colName = n => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const colIndex = ref => { let n = 0; for (const ch of ref.replace(/\d+/g, "")) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
  const dirOf = p => p.slice(0, p.lastIndexOf("/") + 1);
  const relsOf = p => dirOf(p) + "_rels/" + p.slice(p.lastIndexOf("/") + 1) + ".rels";
  const resolve = (base, target) => {
    if (target.startsWith("/")) return target.slice(1);
    const parts = dirOf(base).split("/").filter(Boolean);
    for (const seg of target.split("/")) { if (seg === "..") parts.pop(); else if (seg !== ".") parts.push(seg); }
    return parts.join("/");
  };
  const els = (node, ns, tag) => Array.from(node.getElementsByTagNameNS(ns, tag));

  /* pkg: { zip, parseXML, serialize }  — XML parts are parsed lazily and written back only if touched */
  function Package(zip, parseXML, serialize) {
    const docs = new Map(), dirty = new Set(), created = new Map();
    const P = {
      zip,
      has: p => !!zip.file(p) || created.has(p),
      async doc(p) {
        if (docs.has(p)) return docs.get(p);
        if (created.has(p)) { docs.set(p, parseXML(created.get(p))); return docs.get(p); }
        const f = zip.file(p); if (!f) return null;
        const d = parseXML(await f.async("string")); docs.set(p, d); return d;
      },
      touch(p) { dirty.add(p); },
      create(p, xml) { created.set(p, xml); docs.delete(p); dirty.add(p); },
      async text(p) { const f = zip.file(p); return f ? f.async("string") : null; },
      async flush() {
        for (const p of dirty) {
          const d = docs.get(p);
          if (d) {
            let s = serialize(d);
            if (!/^<\?xml/.test(s) && !p.endsWith(".vml")) s = XML_DECL + s;
            zip.file(p, s, { createFolders: false });
          } else if (created.has(p)) zip.file(p, created.get(p), { createFolders: false });
        }
        dirty.clear();
      },
    };
    return P;
  }

  async function relsDoc(P, part) {
    const rp = relsOf(part);
    if (!P.has(rp)) P.create(rp, XML_DECL + `<Relationships xmlns="${PKG}"></Relationships>`);
    return { path: rp, doc: await P.doc(rp) };
  }
  async function findRel(P, part, type) {
    const { doc } = await relsDoc(P, part);
    const r = els(doc, PKG, "Relationship").find(x => x.getAttribute("Type") === type);
    return r ? { id: r.getAttribute("Id"), target: resolve(part, r.getAttribute("Target")) } : null;
  }
  async function addRel(P, part, type, targetAbs) {
    const { path, doc } = await relsDoc(P, part);
    const ids = els(doc, PKG, "Relationship").map(r => +String(r.getAttribute("Id")).replace(/\D/g, "") || 0);
    const id = "rId" + (Math.max(0, ...ids) + 1);
    const fromDir = dirOf(part).split("/").filter(Boolean), to = targetAbs.split("/");
    let i = 0; while (i < fromDir.length && fromDir[i] === to[i]) i++;
    const rel = "../".repeat(fromDir.length - i) + to.slice(i).join("/");
    const r = doc.createElementNS(PKG, "Relationship");
    r.setAttribute("Id", id); r.setAttribute("Type", type); r.setAttribute("Target", rel);
    doc.documentElement.appendChild(r); P.touch(path);
    return id;
  }
  async function addContentType(P, partAbs, type) {
    const doc = await P.doc("[Content_Types].xml");
    const name = "/" + partAbs;
    if (els(doc, CT, "Override").some(o => o.getAttribute("PartName") === name)) return;
    const o = doc.createElementNS(CT, "Override"); o.setAttribute("PartName", name); o.setAttribute("ContentType", type);
    doc.documentElement.appendChild(o); P.touch("[Content_Types].xml");
  }
  async function ensureVmlDefault(P) {
    const doc = await P.doc("[Content_Types].xml");
    if (els(doc, CT, "Default").some(d => (d.getAttribute("Extension") || "").toLowerCase() === "vml")) return;
    const d = doc.createElementNS(CT, "Default"); d.setAttribute("Extension", "vml"); d.setAttribute("ContentType", CT_VML);
    doc.documentElement.insertBefore(d, doc.documentElement.firstChild); P.touch("[Content_Types].xml");
  }
  function nextFree(P, pattern) { // pattern like "xl/comments#.xml"
    let n = 1; while (P.has(pattern.replace("#", n))) n++; return pattern.replace("#", n);
  }

  // ---------- styles ----------
  function Styles(P) {
    let doc, fills, xfs, cache = new Map();
    return {
      async load() {
        doc = await P.doc("xl/styles.xml");
        fills = doc.getElementsByTagNameNS(MAIN, "fills")[0];
        xfs = doc.getElementsByTagNameNS(MAIN, "cellXfs")[0];
      },
      fillIdOfStyle(s) { const x = els(xfs, MAIN, "xf")[s]; return x ? +(x.getAttribute("fillId") || 0) : 0; },
      solidFill(rgb) {
        const list = els(fills, MAIN, "fill");
        const i = list.findIndex(f => {
          const pf = f.getElementsByTagNameNS(MAIN, "patternFill")[0]; if (!pf || pf.getAttribute("patternType") !== "solid") return false;
          const fg = pf.getElementsByTagNameNS(MAIN, "fgColor")[0];
          return fg && (fg.getAttribute("rgb") || "").toUpperCase() === rgb && !fg.hasAttribute("tint");
        });
        if (i >= 0) return i;
        const f = doc.createElementNS(MAIN, "fill"), pf = doc.createElementNS(MAIN, "patternFill"), fg = doc.createElementNS(MAIN, "fgColor"), bg = doc.createElementNS(MAIN, "bgColor");
        pf.setAttribute("patternType", "solid"); fg.setAttribute("rgb", rgb); bg.setAttribute("indexed", "64");
        pf.appendChild(fg); pf.appendChild(bg); f.appendChild(pf); fills.appendChild(f);
        fills.setAttribute("count", String(list.length + 1)); P.touch("xl/styles.xml");
        return list.length;
      },
      withFill(s, fillId) {  // same style as s, different fill
        const key = s + ":" + fillId; if (cache.has(key)) return cache.get(key);
        const list = els(xfs, MAIN, "xf"); const base = list[s] || list[0];
        if (+(base.getAttribute("fillId") || 0) === fillId) { cache.set(key, s); return s; }
        const x = base.cloneNode(true); x.setAttribute("fillId", String(fillId)); x.setAttribute("applyFill", "1");
        // reuse an identical xf if one already exists
        const want = new XMLSerializerShim().str(x);
        let idx = list.findIndex(e => new XMLSerializerShim().str(e) === want);
        if (idx < 0) { xfs.appendChild(x); idx = list.length; xfs.setAttribute("count", String(idx + 1)); P.touch("xl/styles.xml"); }
        cache.set(key, idx); return idx;
      },
    };
  }
  // tiny attribute-based comparer (avoids depending on a global XMLSerializer in Node)
  function XMLSerializerShim() {}
  XMLSerializerShim.prototype.str = function (el) {
    const a = Array.from(el.attributes).map(x => x.name + "=" + x.value).sort().join("|");
    const k = Array.from(el.childNodes).filter(n => n.nodeType === 1).map(n => this.str(n)).join(";");
    return el.localName + "[" + a + "]{" + k + "}";
  };

  // ---------- sheet cells ----------
  function cellNode(doc, row, col) {
    const sd = doc.getElementsByTagNameNS(MAIN, "sheetData")[0];
    const ref = colName(col) + row;
    let rowEl = els(sd, MAIN, "row").find(r => +r.getAttribute("r") === row);
    if (!rowEl) {
      rowEl = doc.createElementNS(MAIN, "row"); rowEl.setAttribute("r", String(row));
      const after = els(sd, MAIN, "row").find(r => +r.getAttribute("r") > row);
      sd.insertBefore(rowEl, after || null);
    }
    let c = els(rowEl, MAIN, "c").find(x => x.getAttribute("r") === ref);
    if (!c) {
      c = doc.createElementNS(MAIN, "c"); c.setAttribute("r", ref);
      const after = els(rowEl, MAIN, "c").find(x => colIndex(x.getAttribute("r")) > col);
      rowEl.insertBefore(c, after || null);
    }
    return c;
  }
  function clearValue(c) {
    for (const tag of ["v", "f", "is"]) for (const e of els(c, MAIN, tag)) c.removeChild(e);
    c.removeAttribute("t");
  }

  // ---------- threaded comments ----------
  async function personId(P, name) {
    let rel = await findRel(P, "xl/workbook.xml", T.person), path;
    if (rel) path = rel.target;
    else {
      path = "xl/persons/person.xml";
      P.create(path, XML_DECL + `<personList xmlns="${TC_NS}" xmlns:x="${MAIN}"></personList>`);
      await addRel(P, "xl/workbook.xml", T.person, path); await addContentType(P, path, CT_PERSON);
    }
    const doc = await P.doc(path);
    const hit = els(doc, TC_NS, "person").find(p => (p.getAttribute("displayName") || "").trim().toLowerCase() === name.trim().toLowerCase());
    if (hit) return hit.getAttribute("id");
    const p = doc.createElementNS(TC_NS, "person"), id = guid();
    p.setAttribute("displayName", name); p.setAttribute("id", id); p.setAttribute("userId", name); p.setAttribute("providerId", "None");
    doc.documentElement.appendChild(p); P.touch(path);
    return id;
  }

  async function sheetCommentParts(P, sheetPath, sheetDoc) {
    let tc = await findRel(P, sheetPath, T.threaded);
    if (!tc) {
      const p = nextFree(P, "xl/threadedComments/threadedComment#.xml");
      P.create(p, XML_DECL + `<ThreadedComments xmlns="${TC_NS}" xmlns:x="${MAIN}"></ThreadedComments>`);
      await addRel(P, sheetPath, T.threaded, p); await addContentType(P, p, CT_THREADED); tc = { target: p };
    }
    let lc = await findRel(P, sheetPath, T.comments);
    if (!lc) {
      const p = nextFree(P, "xl/comments#.xml");
      P.create(p, XML_DECL + `<comments xmlns="${MAIN}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><authors></authors><commentList></commentList></comments>`);
      await addRel(P, sheetPath, T.comments, p); await addContentType(P, p, CT_COMMENTS); lc = { target: p };
    }
    let vm = await findRel(P, sheetPath, T.vml);
    if (!vm) {
      const p = nextFree(P, "xl/drawings/vmlDrawing#.vml");
      const block = 100 + Math.floor(Math.random() * 800);
      P.create(p, `<xml xmlns:v="urn:schemas-microsoft-com:vml"\r\n xmlns:o="urn:schemas-microsoft-com:office:office"\r\n xmlns:x="urn:schemas-microsoft-com:office:excel">\r\n <o:shapelayout v:ext="edit">\r\n  <o:idmap v:ext="edit" data="${block}"/>\r\n </o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"\r\n  path="m,l,21600r21600,l21600,xe">\r\n  <v:stroke joinstyle="miter"/>\r\n  <v:path gradientshapeok="t" o:connecttype="rect"/>\r\n </v:shapetype></xml>`);
      const rid = await addRel(P, sheetPath, T.vml, p); await ensureVmlDefault(P); vm = { target: p };
      if (!sheetDoc.getElementsByTagNameNS(MAIN, "legacyDrawing").length) {
        const ld = sheetDoc.createElementNS(MAIN, "legacyDrawing"); ld.setAttributeNS(REL_NS, "r:id", rid);
        const ws = sheetDoc.documentElement;
        const before = ["legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]
          .map(t => ws.getElementsByTagNameNS(MAIN, t)[0]).find(Boolean);
        ws.insertBefore(ld, before || null); P.touch(sheetPath);
      }
    }
    return { tc: tc.target, lc: lc.target, vm: vm.target };
  }

  async function addThreadedComment(P, sheetPath, sheetDoc, row, col, text, author, when) {
    const ref = colName(col) + row;
    const parts = await sheetCommentParts(P, sheetPath, sheetDoc);
    const tcDoc = await P.doc(parts.tc);
    if (els(tcDoc, TC_NS, "threadedComment").some(t => t.getAttribute("ref") === ref)) return false; // keep existing thread
    const lcDoc = await P.doc(parts.lc);
    if (els(lcDoc, MAIN, "comment").some(c => c.getAttribute("ref") === ref)) return false;
    const pid = await personId(P, author), id = guid();
    const pad = n => String(n).padStart(2, "0");
    const dT = `${when.getFullYear()}-${pad(when.getMonth() + 1)}-${pad(when.getDate())}T${pad(when.getHours())}:${pad(when.getMinutes())}:${pad(when.getSeconds())}.00`;
    const t = tcDoc.createElementNS(TC_NS, "threadedComment");
    t.setAttribute("ref", ref); t.setAttribute("dT", dT); t.setAttribute("personId", pid); t.setAttribute("id", id);
    const tt = tcDoc.createElementNS(TC_NS, "text"); tt.appendChild(tcDoc.createTextNode(text)); t.appendChild(tt);
    tcDoc.documentElement.appendChild(t); P.touch(parts.tc);
    // legacy fallback note
    const authors = lcDoc.getElementsByTagNameNS(MAIN, "authors")[0];
    const au = lcDoc.createElementNS(MAIN, "author"); au.appendChild(lcDoc.createTextNode("tc=" + id)); authors.appendChild(au);
    const aid = els(authors, MAIN, "author").length - 1;
    const cm = lcDoc.createElementNS(MAIN, "comment");
    cm.setAttribute("ref", ref); cm.setAttribute("authorId", String(aid)); cm.setAttribute("shapeId", "0");
    cm.setAttributeNS("http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "xr:uid", id);
    const ct = lcDoc.createElementNS(MAIN, "text"), ctt = lcDoc.createElementNS(MAIN, "t");
    ctt.appendChild(lcDoc.createTextNode(LEGACY_TEXT + text)); ct.appendChild(ctt); cm.appendChild(ct);
    lcDoc.getElementsByTagNameNS(MAIN, "commentList")[0].appendChild(cm); P.touch(parts.lc);
    // VML shape (string edit: VML is not well-formed XML in general)
    let vml = (await P.text(parts.vm)) || "";
    if (!vml) vml = P._createdText(parts.vm);
    const block = +((/o:idmap[^>]*data="(\d+)/.exec(vml) || [])[1] || 1);
    const ids = [...vml.matchAll(/_x0000_s(\d+)/g)].map(m => +m[1]);
    const sid = ids.length ? Math.max(...ids) + 1 : block * 1024 + 1;
    const shape = `<v:shape id="_x0000_s${sid}" type="#_x0000_t202" style='position:absolute;\r\n  margin-left:0;margin-top:0;width:108pt;height:59.25pt;\r\n  z-index:1;visibility:hidden' fillcolor="infoBackground [80]" strokecolor="none [81]"\r\n  o:insetmode="auto">\r\n  <v:fill color2="infoBackground [80]"/>\r\n  <v:shadow color="none [81]" obscured="t"/>\r\n  <v:path o:connecttype="none"/>\r\n  <v:textbox style='mso-direction-alt:auto'>\r\n   <div style='text-align:left'></div>\r\n  </v:textbox>\r\n  <x:ClientData ObjectType="Note">\r\n   <x:MoveWithCells/>\r\n   <x:SizeWithCells/>\r\n   <x:Anchor>\r\n    ${col}, 15, ${Math.max(0, row - 2)}, 10, ${col + 2}, 15, ${row + 2}, 4</x:Anchor>\r\n   <x:AutoFill>False</x:AutoFill>\r\n   <x:Row>${row - 1}</x:Row>\r\n   <x:Column>${col - 1}</x:Column>\r\n  </x:ClientData>\r\n </v:shape>`;
    vml = vml.replace(/<\/xml>\s*$/, shape + "</xml>");
    P.zip.file(parts.vm, vml, { createFolders: false }); P._forget(parts.vm);
    return true;
  }

  /* edits: [{sheetPath, row, col, action: "book"|"cancel"|"clear", clearFillFrom?: {row,col}}]
     opts: {author, when: Date, comment: true}                                              */
  async function applyEdits(zip, parseXML, serialize, edits, opts) {
    const P = Package(zip, parseXML, serialize);
    P._createdText = () => ""; P._forget = () => {};
    // VML parts created in this session live in "created"; route reads through the zip once flushed
    const createdVml = new Map();
    const origCreate = P.create;
    P.create = (p, xml) => { if (p.endsWith(".vml")) { createdVml.set(p, xml); zip.file(p, xml, { createFolders: false }); } else origCreate(p, xml); };
    P._createdText = p => createdVml.get(p) || "";
    const S = Styles(P); await S.load();
    const report = { cells: 0, comments: 0, sheets: new Set() };
    const bySheet = new Map();
    for (const e of edits) { if (!bySheet.has(e.sheetPath)) bySheet.set(e.sheetPath, []); bySheet.get(e.sheetPath).push(e); }
    for (const [sheetPath, list] of bySheet) {
      const doc = await P.doc(sheetPath);
      const styleOf = (r, c) => +(cellNode(doc, r, c).getAttribute("s") || 0);
      for (const e of list) {
        const c = cellNode(doc, e.row, e.col); const s = +(c.getAttribute("s") || 0);
        let fillId;
        if (e.action === "book") fillId = S.solidFill(FILL.book);
        else if (e.action === "cancel") fillId = S.solidFill(FILL.cancel);
        else { fillId = e.clearFillFrom ? S.fillIdOfStyle(styleOf(e.clearFillFrom.row, e.clearFillFrom.col)) : 0; clearValue(c); }
        c.setAttribute("s", String(S.withFill(s, fillId)));
        report.cells++;
      }
      P.touch(sheetPath); report.sheets.add(sheetPath);
      // one comment per run of consecutive newly booked days in the same row
      if (opts.comment !== false) {
        const booked = list.filter(e => e.action === "book").sort((a, b) => a.row - b.row || a.col - b.col);
        let prev = null;
        for (const e of booked) {
          const start = !prev || prev.row !== e.row || prev.col !== e.col - 1;
          if (start) {
            const d = opts.when; const txt = `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
            if (await addThreadedComment(P, sheetPath, doc, e.row, e.col, txt, opts.author, opts.when)) report.comments++;
          }
          prev = e;
        }
      }
    }
    await P.flush();
    return report;
  }

  root.LeaveEdit = { applyEdits, FILL, colName };
})(typeof window !== "undefined" ? window : globalThis);
