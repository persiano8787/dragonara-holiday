/* Holiday Availability — core reader (shared by manager app and tests).
   Reads the .xlsx package directly (JSZip + XML) so the same model can later be used
   for surgical edits. Rules agreed with Antonio, 16/10/2026:
   - grey on the whole column (STOP LEAVE) = blocked date
   - red = cancelled leave (does not count)
   - any other fill = booking; where there is no fill, the letter decides
   - free places = group limit - bookings; black is only a cross-check
   - hidden / squeezed rows are ignored (people who left)                           */
(function (root) {
  "use strict";
  const MONTHS = ["january","february","march","april","may","june","july","august","september","october","november","december"];
  const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
  const RGB = { FFFFFF00: "Y", FF92D050: "v", FF00B050: "D", FFFF0000: "R", FF0099CC: "B", FFFFC000: "O", FF00B0F0: "c", FFF5A9F7: "p" };
  const NOT_BOOKING = new Set(["K", ".", "w", "s", "G", "g", "R"]);
  const BOOKING_TEXT = /^(L|X|X\/L|L\/X|AL|8)$/i;
  const SKIP_NAME = new RegExp("^(\\d+|20\\d\\d|total|leave request.*|l/request staff|leave company|off|(" + MONTHS.join("|") + ")(\\s*20\\d\\d)?)$", "i");

  const DEPTS = [
    { id: "cashdesk", dept: "Cashdesk", match: /cashdesk/i,
      groups: { senior: "Senior cashier", cashier: "Cashier" }, order: ["Senior cashier", "Cashier"],
      limits: { "Senior cashier": 1, "Cashier": 1 } },
    { id: "fnb", dept: "F&B", match: /f\s*&\s*b|fnb/i,
      groups: { "floor and pantry": "F&B server", "*": "F&B server" }, order: ["F&B server"],
      limits: { "F&B server": "MAX" } },
    { id: "host-reception", dept: "Host & Reception", match: /host/i,
      groups: { "*": "Host & Reception" }, order: ["Host & Reception"],
      limits: { "Host & Reception": 2 } },
    { id: "gaming-poker", dept: "Gaming & Poker", match: /^lg\b|\blg leave/i,
      groups: { inspectors: "Inspectors", "dealer inspectors": "Inspectors", dealers: "Dealers", "poker staff": "Poker dealer" },
      order: ["Inspectors", "Dealers", "Poker supervisor", "Poker dealer"],
      labels: { Inspectors: "Inspectors & Dealer inspectors" },
      limits: { Inspectors: "MAX", Dealers: "MAX", "Poker supervisor": 1, "Poker dealer": 1 },
      supervisors: ["GLODOWSKI", "PEREIRA"] },
    { id: "security-driver", dept: "Security & Driver", match: /sec\s*&\s*driver|security/i,
      groups: { security: "Security", driver: "Driver" }, order: ["Security", "Driver"],
      limits: { Security: 1, Driver: 1 } },
  ];
  const IGNORE_FILE = /gm\s*&\s*pit|pit\s*boss/i;

  function deptForFile(name) {
    if (IGNORE_FILE.test(name)) return null;
    return DEPTS.find(d => d.match.test(name)) || null;
  }

  // ---------- xlsx package helpers ----------
  const kids = (el, tag) => Array.from(el.getElementsByTagNameNS(MAIN, tag));
  const kid = (el, tag) => el.getElementsByTagNameNS(MAIN, tag)[0] || null;
  function colIndex(ref) { let n = 0; for (const ch of ref.replace(/\d+/g, "")) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
  function textOf(el) { return el ? kids(el, "t").map(t => t.textContent).join("") : ""; }

  function fillCode(fill) {
    // mirrors colors.code() of the verified Python reader
    if (!fill) return ".";
    const pf = kid(fill, "patternFill");
    if (!pf) return ".";
    const pt = pf.getAttribute("patternType");
    if (!pt || pt === "none") return ".";
    const fg = kid(pf, "fgColor");
    if (!fg) return pt === "solid" ? "." : "?00";
    if (fg.hasAttribute("rgb")) { const rgb = fg.getAttribute("rgb").toUpperCase(); return RGB[rgb] || "?" + rgb.slice(2, 4); }
    if (fg.hasAttribute("theme")) {
      const th = +fg.getAttribute("theme"); const tint = Math.round((+fg.getAttribute("tint") || 0) * 100) / 100;
      if (th === 0) return ({ 0: ".", "-0.25": "G", "-0.35": "g", "-0.05": "w", "-0.15": "s" })[String(tint)] || "w";
      if (th === 1) return "K";
      return ("t" + th).slice(0, 2);
    }
    return "?";
  }

  async function openWorkbook(JSZip, data, parseXML) {
    const zip = await JSZip.loadAsync(data);
    const xml = async p => { const f = zip.file(p); return f ? parseXML(await f.async("string")) : null; };
    const wb = await xml("xl/workbook.xml");
    const rels = await xml("xl/_rels/workbook.xml.rels");
    const relMap = {};
    for (const r of Array.from(rels.getElementsByTagNameNS(PKG_REL, "Relationship"))) relMap[r.getAttribute("Id")] = r.getAttribute("Target");
    const sst = await xml("xl/sharedStrings.xml");
    const strings = sst ? kids(sst, "si").map(textOf) : [];
    const styles = await xml("xl/styles.xml");
    const fills = kids(kid(styles, "fills"), "fill");
    const xfFill = kids(kid(styles, "cellXfs"), "xf").map(x => +(x.getAttribute("fillId") || 0));
    const codeForStyle = s => fillCode(fills[xfFill[s] || 0]);
    const sheets = kids(kid(wb, "sheets"), "sheet").map(s => {
      let target = relMap[s.getAttributeNS(REL, "id")] || "";
      target = target.startsWith("/") ? target.slice(1) : "xl/" + target.replace(/^\.\//, "");
      return { name: s.getAttribute("name"), path: target };
    });
    async function readSheet(path) {
      const doc = await xml(path); const grid = new Map();
      const hidden = new Set();
      for (const row of kids(doc, "row")) {
        const r = +row.getAttribute("r"); const line = new Map();
        // rows hidden or squeezed to almost zero height (e.g. people who left, kept by copy-paste) are ignored
        if (row.getAttribute("hidden") === "1" || row.getAttribute("hidden") === "true" || (row.hasAttribute("ht") && +row.getAttribute("ht") < 3)) hidden.add(r);
        for (const c of kids(row, "c")) {
          const col = colIndex(c.getAttribute("r")); const s = +(c.getAttribute("s") || 0);
          const t = c.getAttribute("t"); const v = kid(c, "v");
          let val = "";
          if (t === "s" && v) val = strings[+v.textContent] || "";
          else if (t === "inlineStr") val = textOf(kid(c, "is"));
          else if (t === "str" || t === "e") val = v ? v.textContent : "";
          else if (v) val = kid(c, "f") ? "=" : v.textContent;
          line.set(col, { s, code: codeForStyle(s), text: String(val).trim() });
        }
        grid.set(r, line);
      }
      grid.hidden = hidden;
      return grid;
    }
    return { zip, sheets, readSheet };
  }

  function monthKey(name) {
    const m = /^\s*([A-Za-z]+)\D*(20\d\d)/.exec(name);
    if (!m) return null;
    const mi = MONTHS.findIndex(x => x.slice(0, 3) === m[1].toLowerCase().slice(0, 3));
    return mi < 0 ? null : [+m[2], mi + 1];
  }

  // ---------- parse one department workbook ----------
  async function parseDept(cfg, book, start) {
    const out = {}; const issues = []; const roster = {};
    for (const sh of book.sheets) {
      const mk = monthKey(sh.name); if (!mk) continue;
      const [y, mo] = mk; if (y * 100 + mo < start[0] * 100 + start[1]) continue;
      const nd = new Date(y, mo, 0).getDate(); const grid = await book.readSheet(sh.path);
      const cell = (r, c) => (grid.get(r) && grid.get(r).get(c)) || { s: 0, code: ".", text: "" };
      const label = sh.name.trim();
      let dayCol = null;
      for (let r = 1; r <= 5 && !dayCol; r++) for (let c = 1; c <= 3; c++) {
        const v = cell(r, c).text.toLowerCase(), n = cell(r, c + 1).text.toLowerCase();
        if ((v === "1" || v === "1st") && (n === "2" || n === "2nd")) { dayCol = c; break; }
      }
      if (!dayCol) { issues.push({ kind: "structure", month: label, text: "Day header row not found" }); continue; }
      const maxRow = Math.max(0, ...grid.keys());
      let group = cfg.groups["*"] || null; const found = {}; const rows = [];
      for (let r = 1; r <= maxRow; r++) {
        const name = cell(r, 1).text; const key = name.toLowerCase();
        const line = [1, 2, 3].map(c => cell(r, c).text).join(" ").toUpperCase();
        if (key in cfg.groups) group = cfg.groups[key];
        const m = /MAX\.?\s*(\d+)/.exec(line) || /ONLY\s+(\d+)\s+SHOULD/.exec(line);
        if (m && group && !(group in found)) found[group] = +m[1];
        if (!name || key in cfg.groups || SKIP_NAME.test(name) || line.startsWith("ONLY") || line.startsWith("MAX")) continue;
        if (grid.hidden.has(r)) continue;
        let g = group;
        if (cfg.supervisors && group === "Poker dealer" && cfg.supervisors.some(s => name.toUpperCase().includes(s))) g = "Poker supervisor";
        if (g) rows.push({ r, name, group: g });
      }
      const groups = [...new Set(rows.map(x => x.group))].sort();
      const lim = {};
      for (const g of groups) {
        let L = cfg.limits[g]; if (L === "MAX") L = found[g];
        if (L == null) { issues.push({ kind: "structure", month: label, text: `Limit missing for ${g}` }); L = 0; }
        lim[g] = L;
      }
      const res = {}; for (const g of groups) res[g] = { booked: Array(nd).fill(0), black: Array(nd).fill(0), empty: Array(nd).fill(0) };
      const blocked = [];
      const people = rows.map(x => ({ name: x.name, group: x.group, row: x.r, days: [] }));
      for (let d = 1; d <= nd; d++) {
        const col = dayCol + d - 1;
        const cells = rows.map(x => [cell(x.r, col), x.group]);
        cells.forEach(([c], i) => people[i].days.push({ code: c.code, text: c.text }));
        const grey = cells.filter(([c]) => c.code === "G" || c.code === "g").length;
        if (cells.length && grey >= 0.8 * cells.length) { blocked.push(d); continue; }
        for (const [c, g] of cells) {
          if (c.code === "R" || c.code === "G" || c.code === "g") continue;
          if (!NOT_BOOKING.has(c.code) || BOOKING_TEXT.test(c.text)) res[g].booked[d - 1]++;
          else {
            if (c.text && c.text !== ".") issues.push({ kind: "text", month: label, day: d, group: g, text: `Unrecognised text "${c.text}" (not counted)` });
            res[g].empty[d - 1]++;
            if (c.code === "K") res[g].black[d - 1]++;
          }
        }
      }
      const gout = {};
      for (const g of groups) {
        const b = res[g].booked;
        for (let d = 1; d <= nd; d++) {
          if (blocked.includes(d)) continue;
          const e = res[g].empty[d - 1], k = res[g].black[d - 1], left = Math.max(0, lim[g] - b[d - 1]);
          if (e > 0 && k >= 0.8 * e && left > 0) issues.push({ kind: "black", month: label, day: d, group: g, text: `Marked full (black) but ${left} place(s) free (${b[d - 1]}/${lim[g]})` });
          if (b[d - 1] > lim[g]) issues.push({ kind: "over", month: label, day: d, group: g, text: `${b[d - 1]} bookings, limit ${lim[g]}` });
        }
        gout[g] = { limit: lim[g], booked: b };
      }
      const key = `${y}-${String(mo).padStart(2, "0")}`;
      out[key] = { blocked, groups: gout };
      roster[key] = { sheet: sh.name, path: sh.path, dayCol, days: nd, people };
    }
    return { months: out, issues, roster };
  }

  // public data for the staff page: numbers only, never names
  function publicData(cfg, months, updated) {
    return { dept: cfg.dept, updated, groups: cfg.order.map(g => ({ id: g, label: (cfg.labels && cfg.labels[g]) || g })), months };
  }

  root.LeaveCore = { DEPTS, deptForFile, openWorkbook, parseDept, publicData, monthKey };
})(typeof window !== "undefined" ? window : globalThis);
