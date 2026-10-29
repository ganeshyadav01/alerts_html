#!/usr/bin/env node
/*
 * Builds docs/mockups/assets/catalog-data.js from the real workbook, so the mockups show every job
 * and every service. Prototype of the M1 pipeline (DESIGN.md §8–§10): zip/XML reading, merge
 * expansion, sheet-role detection, header aliasing, cell cleaning, service matching, frequency grammar.
 *
 * Usage:  node scripts/build-mockup-data.mjs ["path/to/workbook.xlsx"]
 * Zero dependencies (Node 22). Works on macOS and Windows.
 * The output contains real workbook data: keep it out of git (see DESIGN.md §6).
 */
import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SRC = path.resolve(process.argv.slice(2).find((a) => !a.startsWith("--")) || path.join(ROOT, "raw-inputs", "R+F Jobs Master List.xlsx"));
const OUT = path.join(ROOT, "docs", "mockups", "assets", "catalog-data.js");
const SCHEDULE_TZ = "America/Los_Angeles";

/* ── Zip reader (stored + deflate) ─────────────────────────── */
function readZip(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error("Not a zip/xlsx file");
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const files = new Map();
  for (let n = 0; n < count; n++) {
    const method = buf.readUInt16LE(p + 10), csize = buf.readUInt32LE(p + 20);
    const nlen = buf.readUInt16LE(p + 28), elen = buf.readUInt16LE(p + 30), clen = buf.readUInt16LE(p + 32);
    const local = buf.readUInt32LE(p + 42), name = buf.toString("utf8", p + 46, p + 46 + nlen);
    const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
    const data = buf.subarray(start, start + csize);
    files.set(name, () => (method === 8 ? zlib.inflateRawSync(data) : data).toString("utf8"));
    p += 46 + nlen + elen + clen;
  }
  return (name) => (files.has(name) ? files.get(name)() : null);
}

/* ── XML helpers ───────────────────────────────────────────── */
const decode = (s) => s
  .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
  .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d))
  .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const texts = (xml) => [...xml.replace(/<rPh[\s\S]*?<\/rPh>/g, "").matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((m) => decode(m[1])).join("");
const colIndex = (ref) => { let n = 0; for (const c of ref.match(/[A-Z]+/)[0]) n = n * 26 + c.charCodeAt(0) - 64; return n - 1; };
const colName = (i) => { let s = ""; i++; while (i) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; };
const relTarget = (rels, id) => { const m = rels.match(new RegExp(`<Relationship[^>]*Id="${id}"[^>]*>`)); return m && m[0].match(/Target="([^"]+)"/)[1]; };

function readWorkbook(file) {
  const get = readZip(fs.readFileSync(file));
  const ss = (get("xl/sharedStrings.xml") || "").match(/<si>[\s\S]*?<\/si>/g)?.map(texts) || [];
  const wb = get("xl/workbook.xml"), wbRels = get("xl/_rels/workbook.xml.rels");
  const core = get("docProps/core.xml") || "";
  const modified = (core.match(/<dcterms:modified[^>]*>([^<]+)</) || [])[1] || null;
  const sheets = [...wb.matchAll(/<sheet\s[^>]*>/g)].map((m) => {
    const tag = m[0];
    const name = decode(tag.match(/name="([^"]*)"/)[1]);
    const hidden = /state="(hidden|veryHidden)"/.test(tag);
    const rid = tag.match(/r:id="([^"]+)"/)[1];
    const target = relTarget(wbRels, rid).replace(/^\/?xl\//, "");
    const xml = get("xl/" + target);
    const rows = [], numericCells = [];
    for (const rm of xml.matchAll(/<row\s[^>]*?r="(\d+)"[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const r = +rm[1] - 1, row = [];
      for (const cm of (rm[2] || "").matchAll(/<c\s([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const attrs = cm[1], body = cm[2] || "", ref = attrs.match(/r="([A-Z]+\d+)"/)[1];
        const t = (attrs.match(/\bt="(\w+)"/) || [])[1];
        const v = (body.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
        let val = null;
        if (t === "s" && v != null) val = ss[+v];
        else if (t === "inlineStr") val = texts(body);
        else if (t === "b" && v != null) val = v === "1" ? "TRUE" : "FALSE";
        else if (v != null) { val = decode(v); if (!t) numericCells.push(ref); }
        if (val != null && val !== "") row[colIndex(ref)] = val;
      }
      rows[r] = row;
    }
    let merges = 0;
    for (const mm of xml.matchAll(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g)) {
      merges++;
      const c1 = colIndex(mm[1]), r1 = +mm[2] - 1, c2 = colIndex(mm[3]), r2 = +mm[4] - 1;
      const v = (rows[r1] || [])[c1];
      if (v == null) continue;
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) { rows[r] = rows[r] || []; if (rows[r][c] == null) rows[r][c] = v; }
    }
    // Cell comments (legacy comments part also carries threaded comments as text)
    const comments = {};
    const sRels = get("xl/worksheets/_rels/" + path.posix.basename(target) + ".rels");
    const cRef = sRels && [...sRels.matchAll(/<Relationship[^>]*>/g)].map((x) => x[0]).find((x) => /\/comments"/.test(x));
    if (cRef) {
      const cxml = get(path.posix.normalize("xl/worksheets/" + cRef.match(/Target="([^"]+)"/)[1]));
      for (const c of (cxml || "").matchAll(/<comment\s[^>]*ref="([A-Z]+\d+)"[^>]*>([\s\S]*?)<\/comment>/g)) {
        let txt = texts(c[2]).replace(/\r\n?/g, "\n");
        const i = txt.indexOf("Comment:");
        if (i >= 0) txt = txt.slice(i + 8);
        comments[c[1]] = txt.replace(/\s+/g, " ").trim();
      }
    }
    for (let i = 0; i < rows.length; i++) rows[i] = rows[i] || [];
    return { name, hidden, rows, merges, comments, numericCells };
  });
  return { sheets, modified };
}

/* ── Cleaning and keys (DESIGN.md §8.2) ────────────────────── */
const clean = (v) => { if (v == null) return null; const s = String(v).replace(/\r\n?/g, "\n").replace(/ /g, " ").trim(); return s === "" ? null : s; };
const mkey = (v) => (v == null ? "" : String(v).normalize("NFKC").replace(/\s+/g, " ").trim().toLowerCase());
const hkey = (h) => (h == null ? "" : String(h).toLowerCase().replace(/[^a-z0-9]/g, ""));
const oneLine = (h) => String(h).replace(/\s+/g, " ").trim();

/* ── Header aliases (config/header-aliases.yaml, DESIGN.md §18) ─ */
const FIELDS = [
  ["jobName", ["jobname", "name", "alertname"], []],
  ["jobDescription", ["jobdescription", "description", "alertdescription"], []],
  ["region", ["region", "regions", "market"], []],
  ["jobFrequency", ["jobfrequency", "frequency", "schedule"], []],
  ["state", ["state"], ["stateenabled"]],
  ["monitored", ["monitored"], ["monitoredyes"]],
  ["serviceCategory", ["servicecategory", "category"], []],
  ["businessService", ["businessservice", "service"], []],
  ["alertName", ["alertname", "alertnames", "alertnamedescription"], []],
  ["alertFrequency", ["alertfrequency"], []],
  ["severity", ["severity", "priority", "sev"], []],
  ["resolutionSla", ["resolutionsla", "sla"], []],
  ["l1Team", ["alertresponder"], ["alertresponderl1"]],
  ["l2Team", ["alertresolver"], ["alertresolverl2"]],
  ["stakeholder", ["rfstakeholder", "stakeholder"], []],
  ["systemImpact", ["systemimpact"], []],
  ["businessImpact", ["businessimpact"], []],
  ["impactedUserGroup", ["impactedusergroup"], ["impactedusergroupend"]]
];
function mapHeaders(headerRow) {
  const assigned = {}, attrCols = [];
  headerRow.forEach((h, i) => {
    if (h == null) return;
    const k = hkey(h);
    const f = FIELDS.find(([name, al, pre]) => assigned[name] == null && (al.includes(k) || pre.some((p) => k.startsWith(p))));
    if (f) assigned[f[0]] = i;
    else attrCols.push({ i, label: oneLine(h), duplicateOf: FIELDS.find(([n, al, pre]) => al.includes(k) || pre.some((p) => k.startsWith(p)))?.[0] || null });
  });
  return { assigned, attrCols };
}

/* ── Systems (config/systems.yaml) ─────────────────────────── */
function systemFor(tab) {
  const fixed = { "Hybris_NOC": ["hybris", "Hybris", "HY"], "Solr_NOC": ["solr", "Solr", "SO"], "SF Scheduled Jobs": ["salesforce", "Salesforce", "SF"], "Mulesoft Schedulers": ["mule", "Mule", "MU"] };
  if (fixed[tab]) return { id: fixed[tab][0], name: fixed[tab][1], glyph: fixed[tab][2], cronColumns: [] };
  const m = tab.match(/^D&A - (.+)$/);
  if (m) return { id: "bi", name: "BI / D&A", glyph: "BI", sub: m[1].trim(), cronColumns: ["Schedule interval"] };
  return { id: "other", name: tab, glyph: tab.replace(/[^A-Za-z]/g, "").slice(0, 2).toUpperCase(), cronColumns: [] };
}

/* ── Service mappings (config/service-mappings.yaml, DESIGN.md §9.4) ─ */
const SEED_CATEGORIES = {
  "account management": "Customer Management",
  "supply chain & fulfillment": "Fulfillment and Shipping",
  "inventory management": "Fulfillment and Shipping",
  "finance & payments": "Payments and Billing",
  "consultant compensation": "Commissions and Incentives",
  "product management": "Product and Catalog Management",
  "digital commerce": "Product and Catalog Management",
  "marketing & loyalty": "Loyalty and Rewards",
  "gigya": "Identity and Access Management"
};
const SEED_SERVICES = {
  "customer profile": ["Customer Management", "Profile and Account Maintenance"],
  "data protection": ["Backups and Maintenance", "Data Protection ,ETL Operations"],
  "login & authentication": ["Identity and Access Management", "Login and Authentication"]
};
const STATUS_RANK = { exact: 0, normalized: 1, mapped: 2, inferred: 3, suggested: 4, unmatched: 5 };

/* ── Frequency grammar (DESIGN.md §10.3) ───────────────────── */
const H24 = Array.from({ length: 24 }, (_, i) => i);
const stepArr = (start, end, n) => { const a = []; for (let x = start; x <= end; x += n) a.push(x); return a; };
const range = (a, b) => stepArr(a, b, 1);
const DOWS = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
const DOW_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
const MON_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const ORD = (n) => n + (n % 10 === 1 && n !== 11 ? "st" : n % 10 === 2 && n !== 12 ? "nd" : n % 10 === 3 && n !== 13 ? "rd" : "th");
const pad = (n) => String(n).padStart(2, "0");
const TIME = "(\\d{1,2}(?::\\d{2}){0,2}\\s*(?:am|pm)?)";

function parseTime(s) {
  const m = s.trim().match(/^(\d{1,2})(?::(\d{2}))?(?::(\d{2}))?\s*(am|pm)?$/);
  if (!m) return null;
  let h = +m[1]; const min = +(m[2] || 0), sec = +(m[3] || 0), ap = m[4];
  if (ap === "am" && h === 12) h = 0; else if (ap === "pm" && h < 12) h += 12;
  if (h > 23 || min > 59) return null;
  return { h, m: min, s: sec };
}
function normFreq(raw) {
  return raw.toLowerCase().replace(/\s+/g, " ").trim()
    .replace(/(\d)\s+:(\d)/g, "$1:$2")
    .replace(/\bhrs?\b/g, (x) => (x === "hrs" ? "hours" : "hour"))
    .replace(/\bmins?\b/g, (x) => (x === "mins" ? "minutes" : "minute"));
}
const cal = (o) => Object.assign({ t: "cal" }, o);
const R = (rule, sched, conf, assumptions) => ({ rule, sched, conf, assumptions: assumptions || [] });

function parseCron(expr) {
  const parts = String(expr).trim().split(/\s+/);
  if (parts.length < 5 || parts.length > 6 || !parts.every((p) => /^[\d*\/,\-?LW#a-z]+$/i.test(p))) return null;
  const f = parts.length === 6 ? parts.slice(1) : parts;
  const field = (s, lo, hi) => {
    if (s === "*" || s === "?") return null;
    const out = new Set();
    for (const part of s.split(",")) {
      const m = part.match(/^(\*|\d+)(?:-(\d+))?(?:\/(\d+))?$/);
      if (!m) return undefined;
      const a = m[1] === "*" ? lo : +m[1], b = m[2] != null ? +m[2] : m[1] === "*" || m[3] ? hi : a, st = +(m[3] || 1);
      for (let x = a; x <= b; x += st) out.add(x === 7 && hi === 7 ? 0 : x);
    }
    const arr = [...out].sort((x, y) => x - y);
    return arr.length === hi - lo + 1 || (hi === 7 && arr.length === 7) ? null : arr;
  };
  const m = field(f[0], 0, 59), h = field(f[1], 0, 23), dom = field(f[2], 1, 31), mon = field(f[3], 1, 12), dow = field(f[4], 0, 7);
  if ([m, h, dom, mon, dow].includes(undefined)) return null;
  return cal({ m: m || range(0, 59), h: h || H24, dom: dom || undefined, months: mon || undefined, dow: dow || undefined });
}

function qualifiers(q, s) {
  if (!q) return s;
  q = q.trim();
  let m;
  if (q === "monday through friday") return Object.assign(s, { dow: [1, 2, 3, 4, 5] });
  if ((m = q.match(/^only on (\w+)$/)) && DOWS.includes(m[1])) return Object.assign(s, { dow: [DOWS.indexOf(m[1])] });
  if ((m = q.match(/^on day (\d+) of the month$/))) return Object.assign(s, { dom: [+m[1]] });
  if ((m = q.match(/^between day (\d+) and (\d+) of the month$/))) return Object.assign(s, { dom: range(+m[1], +m[2]) });
  if ((m = q.match(/^on day ([\d\s,throughand]+) of the month$/))) {
    const dom = [];
    for (const seg of m[1].split(/\s+and\s+|,\s*/)) { const r = seg.match(/^(\d+)(?: through (\d+))?$/); if (!r) return null; dom.push(...range(+r[1], +(r[2] || r[1]))); }
    return Object.assign(s, { dom });
  }
  if ((m = q.match(/^every (\d+) days$/))) return Object.assign(s, { dom: stepArr(1, 31, +m[1]) });
  return null;
}

function parseFrequency(raw) {
  if (raw == null) return R("none", { t: "none", reason: "blank" }, "parsed");
  const t = normFreq(raw);
  let m;
  if (/^(manual|manually)$/.test(t)) return R("none", { t: "none", reason: "manual" }, "parsed");
  if (/^(on demand|on-demand|adhoc|ad hoc)$/.test(t)) return R("none", { t: "none", reason: "on demand" }, "parsed");
  if (/^(not scheduled|na|n\/a|none)$/.test(t)) return R("none", { t: "none", reason: "not scheduled" }, "parsed");
  if (/^outbound\b/.test(t)) return R("none", { t: "none", reason: "event" }, "parsed");
  if ((m = raw.trim().match(/^Triggers after successful run of\s+(.+?)\s*$/i))) return R("dependent", { t: "dep", upName: m[1] }, "assumed");
  const cr = parseCron(raw);
  if (cr && /[\d*]/.test(raw) && raw.trim().split(/\s+/).length >= 5 && !/[a-z]{3,}/i.test(raw)) return R("cron", cr, "parsed");
  // runs-at: hours list before the colon, minutes list after
  if ((m = t.match(/^runs at ([\d,]+):([\d,]+)(?:\s*\(day \*, weekday (\d)-(\d)\))?$/))) {
    const s = cal({ h: m[1].split(",").map(Number).sort((a, b) => a - b), m: m[2].split(",").map(Number).sort((a, b) => a - b) });
    if (m[3]) s.dow = range(+m[3], +m[4]);
    return R("runs-at", s, "parsed");
  }
  if ((m = t.match(/^every (?:(\d+) )?minutes?(?: starting at minute (\d+))?(?: \(([\d,]+)\))?$/))) {
    const n = +(m[1] || 1), st = +(m[2] || 0);
    if (m[3]) return R("every-n-minutes", cal({ m: m[3].split(",").map(Number), h: H24 }), "parsed");
    if (60 % n === 0) return R("every-n-minutes", cal({ m: stepArr(st, 59, n), h: H24 }), "parsed");
    return R("every-n-minutes", { t: "int", every: n * 60 }, "assumed", [n + " does not divide 60; treated as a fixed interval from 00:00"]);
  }
  if ((m = t.match(/^every (\d+) seconds?$/))) {
    const n = +m[1];
    if (3600 % n === 0 && n % 60 === 0) return R("every-n-seconds", cal({ m: stepArr(0, 59, n / 60), h: H24 }), "parsed");
    return R("every-n-seconds", { t: "int", every: n }, "assumed", ["Interval not aligned to the clock; anchored at 00:00 Pacific"]);
  }
  if ((m = t.match(/^hourly,? at :?(\d{1,2})(?: minutes past the hour)?$/))) return R("hourly-at", cal({ m: [+m[1]], h: H24 }), "parsed");
  if (/^every hour, starting at midnight\b/.test(t)) return R("hourly-at", cal({ m: [0], h: H24 }), "parsed");
  if ((m = t.match(/^at (\d+) minutes past the hour, every (\d+) hours(.*)$/))) {
    const min = +m[1], n = +m[2], rest = m[3].trim();
    let hours;
    if (!rest) hours = stepArr(0, 23, n);
    else if ((m = rest.match(/^, starting at (.+)$/)) && parseTime(m[1])) hours = stepArr(parseTime(m[1]).h, 23, n);
    else {
      hours = [];
      for (const seg of rest.matchAll(/every (\d+) hours, between (\d{1,2}:\d{2} [ap]m) and (\d{1,2}:\d{2} [ap]m)/g)) {
        hours.push(...stepArr(parseTime(seg[2]).h, parseTime(seg[3]).h, +seg[1]));
      }
      const first = rest.match(/^, between (\d{1,2}:\d{2} [ap]m) and (\d{1,2}:\d{2} [ap]m)/);
      if (first) hours.push(...stepArr(parseTime(first[1]).h, parseTime(first[2]).h, n));
      if (!hours.length) return R("unparseable", { t: "none", reason: "unparseable" }, "unparseable");
      hours = [...new Set(hours)].sort((a, b) => a - b);
    }
    return R("cronstrue", cal({ m: [min], h: hours }), "parsed");
  }
  if ((m = t.match(new RegExp("^(?:daily |everyday |every day )?at " + TIME + "(?: \\(midnight\\))?(?:, (.+))?$")))) {
    const tm = parseTime(m[1]);
    if (tm) { const s = qualifiers(m[2], cal({ m: [tm.m], h: [tm.h] })); if (s) return R("at-time", s, "parsed"); }
  }
  if ((m = t.match(new RegExp("^twice daily at " + TIME + " and " + TIME + "$")))) {
    const a = parseTime(m[1]), b = parseTime(m[2]);
    if (a && b && a.m === b.m) return R("twice-daily", cal({ m: [a.m], h: [a.h, b.h].sort((x, y) => x - y) }), "parsed");
  }
  if ((m = t.match(new RegExp("^every (" + DOWS.join("|") + ") at " + TIME + "$")))) {
    const tm = parseTime(m[2]); if (tm) return R("weekly", cal({ m: [tm.m], h: [tm.h], dow: [DOWS.indexOf(m[1])] }), "parsed");
  }
  if ((m = t.match(new RegExp("^(?:monthly on the (\\d+)(?:st|nd|rd|th) day|(\\d+)(?:st|nd|rd|th) day of the month) at " + TIME + "$")))) {
    const tm = parseTime(m[3]); if (tm) return R("monthly", cal({ m: [tm.m], h: [tm.h], dom: [+(m[1] || m[2])] }), "parsed");
  }
  if ((m = t.match(new RegExp("^" + TIME + " on the (\\d)(?:st|nd|rd|th) (" + DOWS.join("|") + ") of (.+)$")))) {
    const tm = parseTime(m[1]), months = m[4].replace(/\band\b/g, ",").split(",").map((x) => MONTHS.indexOf(x.trim()) + 1).filter((x) => x > 0);
    if (tm && months.length) return R("nth-weekday", cal({ m: [tm.m], h: [tm.h], sec: tm.s, nth: { wd: DOWS.indexOf(m[3]), n: +m[2] }, months }), "parsed");
  }
  if ((m = t.match(/^(daily|weekly|monthly) - every (\d+)(?: - every (\d+) (minutes|hours))?$/))) {
    const n = +m[2], every = m[3] ? +m[3] : null, unit = m[4];
    const as = ["Informatica-style text; start time assumed 00:00"];
    let s = every == null ? cal({ m: [0], h: [0] }) : unit === "minutes" ? (60 % every === 0 ? cal({ m: stepArr(0, 59, every), h: H24 }) : { t: "int", every: every * 60 }) : (24 % every === 0 ? cal({ m: [0], h: stepArr(0, 23, every) }) : { t: "int", every: every * 3600 });
    if (m[1] === "weekly") { s.dow = [1]; as.push("Weekday not given; assumed Monday"); }
    if (m[1] === "monthly") { s.dom = [Math.max(1, n)]; as.push("Read the number as day " + Math.max(1, n) + " of the month"); }
    return R("informatica", s, "assumed", as);
  }
  if ((m = t.match(/^every (?:(\d+) )?hours?$/))) {
    const n = +(m[1] || 1);
    if (!m[1]) return R("every-n-hours", cal({ m: [0], h: H24 }), "parsed");
    if (n === 24) return R("every-n-hours", cal({ m: [0], h: [0] }), "assumed", ["Start time not given; assumed 00:00"]);
    if (24 % n === 0) return R("every-n-hours", cal({ m: [0], h: stepArr(0, 23, n) }), "assumed", ["Start not given; assumed 00:00, minute 0"]);
    return R("every-n-hours", { t: "int", every: n * 3600 }, "assumed", [n + " does not divide 24; treated as a fixed interval from 00:00"]);
  }
  if ((m = t.match(/^every (\d+) days$/))) return R("every-n-days", cal({ m: [0], h: [0], dstep: +m[1] }), "assumed", ["Start date unknown; anchored 2026-01-01", "Time not given; assumed 00:00"]);
  if (/^(daily|every day|everyday)$/.test(t)) return R("daily-no-time", cal({ m: [0], h: [0] }), "assumed", ["Time not given; assumed 00:00"]);
  return R("unparseable", { t: "none", reason: "unparseable" }, "unparseable");
}

function listDays(a) {
  if (a.length > 2 && a.every((x, i) => i === 0 || x === a[i - 1] + 1)) return a[0] + "–" + a[a.length - 1];
  return a.join(", ");
}
function humanOf(s) {
  if (s.t === "none") return { "blank": "No frequency given", "manual": "Manual", "on demand": "On demand", "not scheduled": "Not scheduled", "event": "Event-driven (no schedule)", "unparseable": "Unknown frequency", "upstream not found": "Upstream job not found" }[s.reason] || s.reason;
  if (s.t === "int") { const sec = s.every; return sec % 3600 === 0 ? "Every " + sec / 3600 + " h" : sec % 60 === 0 ? "Every " + sec / 60 + " min" : "Every " + Math.floor(sec / 60) + " min " + (sec % 60) + " s"; }
  let day = "Daily";
  if (s.nth) day = ORD(s.nth.n) + " " + DOW_SHORT[s.nth.wd] + " of " + s.months.map((x) => MON_SHORT[x - 1]).join(", ");
  else if (s.dom) day = (s.dom.length === 1 ? "Day " : "Days ") + listDays(s.dom) + (s.months ? " of " + s.months.map((x) => MON_SHORT[x - 1]).join(", ") : "");
  else if (s.dow) day = s.dow.length === 5 && s.dow.join() === "1,2,3,4,5" ? "Mon–Fri" : s.dow.map((d) => DOW_SHORT[d]).join(", ");
  else if (s.dstep) day = "Every " + s.dstep + " days";
  let time;
  if (s.h.length === 24) {
    if (s.m.length === 1) time = "hourly at :" + pad(s.m[0]);
    else { const d = s.m[1] - s.m[0]; const even = s.m.every((x, i) => i === 0 || x - s.m[i - 1] === d) && 60 % d === 0 && s.m.length === 60 / d; time = even ? "every " + d + " min" + (s.m[0] ? " from :" + pad(s.m[0]) : "") : "at :" + s.m.map(pad).join(", :") + " each hour"; }
  } else {
    const ts = []; s.h.forEach((h) => s.m.forEach((mm) => ts.push(pad(h) + ":" + pad(mm) + (s.sec ? ":" + pad(s.sec) : ""))));
    const dh = s.h[1] - s.h[0], evenH = s.h.length > 2 && s.h.every((x, i) => i === 0 || x - s.h[i - 1] === dh);
    time = ts.length <= 6 ? ts.join(", ") : evenH && s.m.length === 1 ? "every " + dh + " h from " + ts[0] : ts.length + " runs (" + ts[0] + "–" + ts[ts.length - 1] + ")";
  }
  if (day === "Daily" && (s.h.length === 24 || /^every/.test(time))) return time[0].toUpperCase() + time.slice(1);
  return day + (s.h.length === 24 ? ", " : " ") + time;
}
const isDailySched = (s) => s.t === "int" || (s.t === "cal" && !s.dom && !s.dow && !s.months && !s.nth && !s.dstep);

/* ── Build ─────────────────────────────────────────────────── */
const { sheets, modified } = readWorkbook(SRC);
const stat = fs.statSync(SRC);
const report = { tabs: [], skipped: [], issues: {}, rules: {}, totals: {} };
function issue(code, level, message, example) {
  const it = report.issues[code] || (report.issues[code] = { code, level, message, count: 0, examples: [] });
  it.count++; if (example && it.examples.length < 6) it.examples.push(example);
}

const roleOf = (sh) => {
  let hi = -1;
  for (let r = 0; r < Math.min(10, sh.rows.length); r++) if (sh.rows[r].filter((v) => clean(v) != null).length >= 3) { hi = r; break; }
  if (hi < 0) return { role: "ignored", hi };
  const { assigned } = mapHeaders(sh.rows[hi]);
  if (assigned.serviceCategory != null && assigned.businessService != null && assigned.jobDescription != null && assigned.jobName == null) return { role: "master", hi };
  if (assigned.jobName != null && assigned.jobFrequency != null && (assigned.state != null || assigned.monitored != null || assigned.serviceCategory != null)) return { role: "jobs", hi };
  return { role: "ignored", hi };
};

// Master
const master = []; // [catName, [[svc, desc]]]
const rawName = new Map(); // master entry -> untrimmed cell text, for "exact" matching
let masterSheet = null;
for (const sh of sheets) {
  const { role, hi } = roleOf(sh);
  sh.role = role; sh.hi = hi;
  if (role === "master" && !masterSheet) masterSheet = sh;
}
{
  const sh = masterSheet, { assigned } = mapHeaders(sh.rows[sh.hi]);
  let cur = null, svcCount = 0;
  for (let r = sh.hi + 1; r < sh.rows.length; r++) {
    const row = sh.rows[r], c = clean(row[assigned.serviceCategory]), s = clean(row[assigned.businessService]), d = clean(row[assigned.jobDescription]);
    if (!c && !s && !d) continue;
    if (c) { cur = [c, []]; rawName.set(cur, String(row[assigned.serviceCategory])); master.push(cur); if (!s) continue; }
    if (!cur) { issue("MASTER_ORPHAN_ROW", "warning", "Service row before any category row.", sh.name + " r" + (r + 1)); continue; }
    if (!s) { issue("MASTER_ORPHAN_ROW", "warning", "Row has a description but no service.", sh.name + " r" + (r + 1)); continue; }
    if (cur[1].some((x) => mkey(x[0]) === mkey(s))) { issue("MASTER_DUPLICATE_SERVICE", "warning", "Duplicate service in the same category; first one kept.", sh.name + " r" + (r + 1) + ' "' + s + '"'); continue; }
    const entry = [s, d || ""]; rawName.set(entry, String(row[assigned.businessService])); cur[1].push(entry); svcCount++;
  }
  report.tabs.push({ name: sh.name, role: "master", rows: svcCount, categories: master.length, merges: sh.merges, extra: [], warnings: 0 });
}
const catByKey = new Map(master.map((c) => [mkey(c[0]), c]));
const svcIndex = new Map(); // svcKey -> [[cat, svc]]
master.forEach((c) => c[1].forEach((s) => { const k = mkey(s[0]); if (!svcIndex.has(k)) svcIndex.set(k, []); svcIndex.get(k).push([c[0], s[0]]); }));

// Jobs
const jobs = [], systems = {}, seen = new Map();
for (const sh of sheets) {
  if (sh === masterSheet) continue;
  if (sh.hidden) { report.skipped.push({ name: sh.name, reason: sh.role === "jobs" ? "hidden" : "hidden · not a job tab", rows: sh.rows.slice(Math.max(sh.hi, 0) + 1).filter((r) => r.some((v) => clean(v) != null)).length }); continue; }
  if (sh.role !== "jobs") { report.skipped.push({ name: sh.name, reason: "not recognised as a job tab", rows: 0 }); continue; }
  const sys = systemFor(sh.name);
  systems[sys.id] = { name: sys.name, glyph: sys.glyph, tab: sys.id === "bi" ? "D&A - *" : sh.name };
  const header = sh.rows[sh.hi], { assigned, attrCols } = mapHeaders(header);
  const extra = attrCols.map((a) => (a.duplicateOf ? a.label + " (col " + colName(a.i) + ")" : a.label));
  attrCols.filter((a) => a.duplicateOf).forEach((a) => issue("HEADER_DUPLICATE", "info", "Second column for an already-mapped field kept as an additional attribute.", sh.name + ': "' + a.label + '" col ' + colName(a.i)));
  FIELDS.forEach(([f]) => { if (assigned[f] == null) issue("HEADER_MISSING", "info", "Expected column not present in this tab; values read as blank.", sh.name + ": " + f); });
  const headerWidth = header.length;
  let rowsParsed = 0, tabWarnings = 0;
  const before = Object.fromEntries(Object.entries(report.issues).map(([k, v]) => [k, v.count]));
  for (let r = sh.hi + 1; r < sh.rows.length; r++) {
    const row = sh.rows[r];
    if (!row.some((v) => clean(v) != null)) continue;
    const g = (f) => (assigned[f] == null ? null : clean(row[assigned[f]]));
    let name = g("jobName");
    const rowNo = r + 1, where = sh.name + " r" + rowNo;
    if (!name) { name = "(unnamed row " + rowNo + ")"; issue("JOB_NAME_MISSING", "warning", "Row has data but no job name; kept as an unnamed job.", where); }
    rowsParsed++;
    // attributes: every unmapped header column, plus blank-header columns that hold data
    const attrs = {};
    attrCols.forEach((a) => { attrs[a.duplicateOf ? a.label + " (col " + colName(a.i) + ")" : a.label] = clean(row[a.i]); });
    for (let c = 0; c < row.length; c++) {
      if (c < headerWidth && header[c] != null) continue;
      const v = clean(row[c]);
      if (v != null) { attrs["Column " + colName(c)] = v; issue("HEADER_BLANK_WITH_DATA", "info", "Cell holds data under a blank header; kept as an additional attribute.", where + " col " + colName(c)); }
    }
    // comments on this row
    Object.keys(sh.comments).forEach((ref) => {
      if (+ref.match(/\d+/)[0] !== rowNo) return;
      const c = colIndex(ref), label = header[c] != null ? oneLine(header[c]) : "Column " + colName(c);
      attrs["Note on " + label] = sh.comments[ref];
      issue("CELL_COMMENT", "info", "Cell comment kept as an additional attribute.", where + " " + ref);
    });

    // State / monitored / severity
    const stateRaw = g("state"), monRaw = g("monitored"), sevRaw = g("severity");
    const sk = mkey(stateRaw);
    const state = ["enabled", "active", "on"].includes(sk) ? "Enabled" : ["disabled", "inactive", "paused", "off"].includes(sk) ? "Disabled" : "Unknown";
    if (state === "Unknown") issue("STATE_BLANK", "info", "State is blank or not recognised; shown as Unknown.", where + (stateRaw ? ' "' + stateRaw + '"' : ""));
    const mk = mkey(monRaw);
    const mon = ["yes", "y", "true", "enabled"].includes(mk) ? true : ["no", "n", "false"].includes(mk) ? false : null;
    if (mon == null) issue("MONITORED_BLANK", "info", "Monitored is blank or not recognised; shown as Unknown.", where + (monRaw ? ' "' + monRaw + '"' : ""));
    const sv = mkey(sevRaw).replace(/\s+/g, "");
    const sev = ["sev1", "s1", "p1", "critical", "high"].includes(sv) ? "Sev1" : ["sev2", "s2", "p2", "medium"].includes(sv) ? "Sev2" : ["sev3", "s3", "p3", "low"].includes(sv) ? "Sev3" : "Unknown";
    if (!sevRaw) issue("SEVERITY_BLANK", "info", "Severity is blank.", where);
    ["resolutionSla", "l1Team", "l2Team"].forEach((f) => { if (!g(f)) issue(f === "resolutionSla" ? "SLA_BLANK" : f === "l1Team" ? "L1_BLANK" : "L2_BLANK", "info", (f === "resolutionSla" ? "Resolution SLA" : f === "l1Team" ? "Alert Responder (L1)" : "Alert Resolver (L2)") + " is blank.", where); });

    // Region tokens
    const regionRaw = g("region");
    (regionRaw || "").split(/[,&/;()]|\band\b/i).map((x) => x.trim().toUpperCase()).filter(Boolean).forEach((tok) => {
      if (!["US", "CA", "AU", "JP", "GLOBAL", "ALL"].includes(tok)) issue("REGION_UNKNOWN", "info", "Region token not in config/regions.yaml.", where + ' "' + tok + '"');
    });
    if (!regionRaw) issue("REGION_BLANK", "info", "Region is blank.", where);

    // Alerts
    const alertsRaw = g("alertName");
    const SENT = ["no alerts configured for this api", "na", "n/a", "none", "-"];
    let alerts = [];
    if (alertsRaw && !SENT.includes(mkey(alertsRaw))) alerts = alertsRaw.split(/\n|,/).map((x) => x.trim()).filter(Boolean);

    // Service matching (DESIGN.md §9.1)
    const catRaw = g("serviceCategory"), svcRaw = g("businessService");
    // exact = the untouched cell text equals the master text (trailing / double spaces make it "normalized")
    const rawOf = (f) => (assigned[f] == null || row[assigned[f]] == null ? null : String(row[assigned[f]]).replace(/\r\n?/g, "\n"));
    const catCell = rawOf("serviceCategory"), svcCell = rawOf("businessService");
    let cat = null, catSt = "unmatched";
    if (catRaw) {
      const exact = master.find((c) => rawName.get(c) === catCell);
      if (exact) { cat = exact; catSt = "exact"; }
      else if (catByKey.has(mkey(catRaw))) { cat = catByKey.get(mkey(catRaw)); catSt = "normalized"; }
      else if (SEED_CATEGORIES[mkey(catRaw)]) { cat = catByKey.get(mkey(SEED_CATEGORIES[mkey(catRaw)])); catSt = "suggested"; }
    }
    let svc = null, svcSt = "unmatched";
    if (svcRaw) {
      const pool = cat ? cat[1] : [];
      const ex = pool.find((s) => rawName.get(s) === svcCell), nm = pool.find((s) => mkey(s[0]) === mkey(svcRaw));
      if (ex) { svc = ex[0]; svcSt = "exact"; }
      else if (nm) { svc = nm[0]; svcSt = "normalized"; }
      else if (SEED_SERVICES[mkey(svcRaw)] && (!cat || cat[0] === SEED_SERVICES[mkey(svcRaw)][0])) {
        const [tc, ts] = SEED_SERVICES[mkey(svcRaw)];
        if (!cat) { cat = catByKey.get(mkey(tc)); catSt = "suggested"; }
        svc = cat[1].find((s) => mkey(s[0]) === mkey(ts))[0]; svcSt = "suggested";
      } else if (!cat && (svcIndex.get(mkey(svcRaw)) || []).length === 1) {
        const [tc, ts] = svcIndex.get(mkey(svcRaw))[0];
        cat = catByKey.get(mkey(tc)); catSt = "inferred"; svc = ts; svcSt = "inferred";
      }
    }
    const status = [catSt, svcSt].sort((a, b) => STATUS_RANK[b] - STATUS_RANK[a])[0];
    if (!catRaw) issue("CATEGORY_BLANK", "warning", "Service Category is blank; job placed under Unmapped (unless the service identifies it).", where);
    else if (!cat) issue("CATEGORY_UNMATCHED", "warning", "Service Category not found in the master tab; job placed under Unmapped.", where + ' "' + catRaw + '"');
    if (cat && !svc) issue(svcRaw ? "SERVICE_UNMATCHED" : "SERVICE_BLANK", "warning", svcRaw ? "Category matched but Business Service not found in it; job placed under Unassigned service." : "Business Service is blank; job placed under Unassigned service.", where + (svcRaw ? ' "' + svcRaw + '"' : ""));
    if (catSt === "suggested" || svcSt === "suggested") issue("MAPPING_SUGGESTED", "info", "Placed through a suggested mapping that still needs confirmation.", where);
    if (catSt === "inferred") issue("CATEGORY_INFERRED", "info", "Category inferred from a service name that is unique in the master tab.", where);

    // Frequency, with optional cron column
    const freq = g("jobFrequency");
    let fr = parseFrequency(freq), source = "frequency", cronCol = null, conflict = null;
    for (const cc of sys.cronColumns) {
      const k = Object.keys(attrs).find((a) => mkey(a) === mkey(cc));
      const val = k && attrs[k];
      const parsed = val && parseCron(val);
      if (parsed) {
        if (fr.sched.t === "none" || fr.sched.t === "dep") {
          conflict = "Text says no fixed schedule; cron column has one";
          issue("FREQ_CONFLICT", "warning", "Frequency text says no fixed schedule (manual / on demand / triggered) but the cron column has one; cron column used.", where + ' "' + freq + '" vs "' + val + '"');
        } else if (fr.conf === "parsed" && fr.sched.t === "cal" && JSON.stringify([fr.sched.m, fr.sched.h, fr.sched.dom, fr.sched.dow]) !== JSON.stringify([parsed.m, parsed.h, parsed.dom, parsed.dow]))
          { conflict = "Text and cron column give different times"; issue("FREQ_CONFLICT_TIME", "warning", "Frequency text and cron column give different times; cron column used.", where + ' "' + freq + '" vs "' + val + '"'); }
        fr = R("cron (column)", parsed, "parsed"); fr.textRule = parseFrequency(freq); source = "cronColumn"; cronCol = val;
      } else if (val) issue("CRON_INVALID", "info", "Cron column value is not a valid cron expression; ignored.", where + ' "' + val + '"');
    }

    const dupKey = sh.name + "|" + mkey(name) + "|" + mkey(regionRaw);
    const dupN = (seen.get(dupKey) || 0) + 1; seen.set(dupKey, dupN);
    if (dupN > 1) issue("DUPLICATE_JOB", "info", "Same job name and region appear more than once in this tab; kept as separate jobs.", where + ' "' + name + '"');
    const id = sys.id + "-" + (sys.sub ? sys.sub.replace(/[^A-Za-z]/g, "").slice(0, 4).toLowerCase() + "-" : "") + "r" + rowNo;

    jobs.push({
      id, sys: sys.id, sub: sys.sub, tab: sh.name, row: rowNo, name, desc: g("jobDescription"), region: regionRaw,
      state, stateRaw, mon, monRaw, catRaw, svcRaw, status, place: cat ? { cat: cat[0], svc: svc || null } : { cat: null },
      alerts, alertsRaw, alertFreq: g("alertFrequency"), sev, sevRaw, sla: g("resolutionSla"), l1: g("l1Team"), l2: g("l2Team"),
      owner: g("stakeholder"), sysImpact: g("systemImpact"), bizImpact: g("businessImpact"), group: g("impactedUserGroup"),
      freq, cronCol, source, textHuman: fr.textRule ? (fr.textRule.sched.t === "dep" ? "After " + fr.textRule.sched.upName : humanOf(fr.textRule.sched)) : undefined, conflict: conflict || undefined, rule: fr.rule, conf: fr.conf, assumptions: fr.assumptions.length ? fr.assumptions : undefined, sched: fr.sched,
      attrs: Object.keys(attrs).length ? attrs : undefined
    });
  }
  for (const [k, v] of Object.entries(report.issues)) if (v.level === "warning") tabWarnings += v.count - (before[k] || 0);
  report.tabs.push({ name: sh.name, role: "jobs", sys: sys.id, sub: sys.sub, rows: rowsParsed, merges: sh.merges, extra, warnings: tabWarnings });
}

// Dependent jobs (DESIGN.md §10.5)
const baseKey = (s) => mkey(s).replace(/\s+dag$/, "").replace(/[-_]prod$/, "");
const byName = new Map();
jobs.forEach((j) => { const k = baseKey(j.name); if (!byName.has(k)) byName.set(k, []); byName.get(k).push(j); });
function resolveDep(j, depth) {
  if (j.sched.t !== "dep") return j;
  if (depth > 8) return null;
  const cands = byName.get(baseKey(j.sched.upName)) || [];
  const up = cands.find((c) => c.tab === j.tab && c !== j) || cands.find((c) => c !== j);
  if (!up) return null;
  return resolveDep(up, depth + 1) && up;
}
jobs.forEach((j) => {
  if (j.sched.t !== "dep") return;
  const upName = j.sched.upName, up = resolveDep(j, 0);
  let root = up; for (let i = 0; root && root.sched.t === "dep" && i < 9; i++) root = resolveDep(root, 0);
  if (up && root && root.sched.t !== "none") {
    j.sched = { t: "dep", up: up.id }; j.assumptions = ["Inherits the schedule of " + up.name, "Offset 0 min"];
    j.human = "After " + up.name + " (≈ " + humanOf(root.sched) + ")";
  } else {
    j.sched = { t: "none", reason: "upstream not found" }; j.conf = "parsed";
    j.human = "After " + upName + " (upstream " + (up ? "has no schedule" : "not found") + ")";
    issue("FREQ_DEPENDENCY_UNRESOLVED", "warning", "Upstream job of a 'Triggers after…' frequency not found or not scheduled; listed as Unscheduled.", j.tab + " r" + j.row + ' "' + upName + '"');
  }
});
jobs.forEach((j) => {
  if (!j.human) j.human = humanOf(j.sched);
  if (j.conf === "unparseable") issue("FREQ_UNPARSEABLE", "warning", "Job Frequency could not be interpreted; listed as Unscheduled.", j.tab + " r" + j.row + ' "' + (j.freq || "") + '"');
  if (j.conf === "assumed" && j.sched.t !== "dep") issue("FREQ_ASSUMED", "info", "Schedule derived with defaults (see the job drawer for assumptions).", j.tab + " r" + j.row + ' "' + j.freq + '"');
  const rk = j.rule + "|" + j.conf;
  report.rules[rk] = (report.rules[rk] || 0) + 1;
});

// Alert-name variants (_Not_Finished vs _NotFinished)
{
  const byKey = new Map();
  jobs.forEach((j) => j.alerts.forEach((a) => { const k = a.toLowerCase().replace(/[_\-\s]/g, ""); if (!byKey.has(k)) byKey.set(k, new Set()); byKey.get(k).add(a); }));
  const suffix = new Map();
  jobs.forEach((j) => j.alerts.forEach((a) => { const m = a.match(/_(Not_?Finished|Not_?Started|Not_?Completed)$/i); if (m) { const k = m[1].toLowerCase().replace("_", ""); if (!suffix.has(k)) suffix.set(k, new Set()); suffix.get(k).add(m[1]); } }));
  suffix.forEach((set) => { if (set.size > 1) jobs.forEach((j) => j.alerts.forEach((a) => { if (/_Not_(Finished|Started|Completed)$/i.test(a)) issue("ALERT_NAME_VARIANT", "info", "Alert suffix spelled with an extra underscore (e.g. _Not_Finished vs _NotFinished).", j.tab + " r" + j.row + " " + a); })); });
  byKey.forEach((set) => { if (set.size > 1) issue("ALERT_NAME_DUPLICATE_VARIANT", "info", "Same alert written differently in different rows.", [...set].join(" / ")); });
}

// Mapping helper rows: every distinct unmatched / suggested raw value
const tokens = (s) => new Set(mkey(s).split(/[^a-z0-9]+/).filter((w) => w && !["and", "or", "the", "of", "a"].includes(w)));
const jac = (a, b) => { const A = tokens(a), B = tokens(b); let i = 0; A.forEach((x) => { if (B.has(x)) i++; }); return i / (A.size + B.size - i || 1); };
const mapping = new Map();
jobs.forEach((j) => {
  const catKnown = j.catRaw && (master.some((c) => c[0] === j.catRaw) || catByKey.has(mkey(j.catRaw)));
  if (j.catRaw && !catKnown) {
    const k = "Category|" + mkey(j.catRaw);
    const e = mapping.get(k) || { kind: "Category", raw: j.catRaw.trim(), jobs: 0, to: SEED_CATEGORIES[mkey(j.catRaw)] || null };
    if (!e.to) { const best = master.map((c) => [c[0], jac(j.catRaw, c[0])]).sort((a, b) => b[1] - a[1])[0]; if (best[1] >= 0.34) e.hint = best[0]; }
    e.st = e.to ? "suggested" : "none"; e.jobs++; mapping.set(k, e);
  }
  if (j.svcRaw && j.place.cat && (j.place.svc == null || ["suggested"].includes(j.status) && SEED_SERVICES[mkey(j.svcRaw)])) {
    const k = "Service|" + j.place.cat + "|" + mkey(j.svcRaw);
    const seed = SEED_SERVICES[mkey(j.svcRaw)];
    const e = mapping.get(k) || { kind: "Service", raw: j.svcRaw.trim(), within: j.place.cat, jobs: 0, to: seed ? seed[0] + " / " + seed[1] : null };
    if (!e.to) { const best = master.find((c) => c[0] === j.place.cat)[1].map((s) => [s[0], jac(j.svcRaw, s[0])]).sort((a, b) => b[1] - a[1])[0]; if (best && best[1] >= 0.34) e.hint = j.place.cat + " / " + best[0]; }
    e.st = e.to ? "suggested" : "none"; e.jobs++; mapping.set(k, e);
  }
});

report.mapping = [...mapping.values()].sort((a, b) => (a.st === "suggested" ? 0 : 1) - (b.st === "suggested" ? 0 : 1) || b.jobs - a.jobs);
report.issues = Object.values(report.issues).sort((a, b) => (a.level === "warning" ? 0 : 1) - (b.level === "warning" ? 0 : 1) || b.count - a.count);
report.rules = Object.entries(report.rules).map(([k, n]) => { const [rule, conf] = k.split("|"); return { rule, conf, count: n }; }).sort((a, b) => ["parsed", "assumed", "unparseable"].indexOf(a.conf) - ["parsed", "assumed", "unparseable"].indexOf(b.conf) || b.count - a.count);
const services = master.reduce((n, c) => n + c[1].length, 0);
report.totals = {
  jobTabs: report.tabs.filter((t) => t.role === "jobs").length, skipped: report.skipped.length, jobs: jobs.length, categories: master.length, services,
  unmapped: jobs.filter((j) => !j.place.cat).length, unassigned: jobs.filter((j) => j.place.cat && !j.place.svc).length,
  fullyMatched: jobs.filter((j) => j.place.cat && j.place.svc && ["exact", "normalized"].includes(j.status)).length,
  suggested: jobs.filter((j) => j.status === "suggested").length, inferred: jobs.filter((j) => j.status === "inferred").length,
  parsed: jobs.filter((j) => j.conf === "parsed").length, assumed: jobs.filter((j) => j.conf === "assumed").length, unparseable: jobs.filter((j) => j.conf === "unparseable").length,
  unscheduled: jobs.filter((j) => j.sched.t === "none").length, numericCells: sheets.filter((s) => !s.hidden && s.role === "jobs").reduce((n, s) => n + s.numericCells.length, 0)
};

const fmtPT = (d) => {
  const date = new Intl.DateTimeFormat("en-GB", { timeZone: SCHEDULE_TZ, day: "numeric", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit", hourCycle: "h23" }).format(d).replace(",", "");
  const zone = new Intl.DateTimeFormat("en-US", { timeZone: SCHEDULE_TZ, timeZoneName: "short" }).formatToParts(d).find((x) => x.type === "timeZoneName").value;
  return date + " " + zone;
};
const now = new Date();
const todayPT = new Intl.DateTimeFormat("en-CA", { timeZone: SCHEDULE_TZ, year: "numeric", month: "2-digit", day: "2-digit" }).format(now);
const nowPT = new Intl.DateTimeFormat("en-GB", { timeZone: SCHEDULE_TZ, hour: "2-digit", minute: "2-digit", hourCycle: "h23" }).format(now).split(":").map(Number);
const relPath = path.relative(ROOT, SRC).split(path.sep).join("/");

const data = {
  meta: {
    workbook: path.basename(SRC), workbookPath: relPath.startsWith("..") ? path.basename(SRC) : relPath,
    modified: fmtPT(modified ? new Date(modified) : stat.mtime), loadedAt: fmtPT(now), generatedAt: now.toISOString(),
    version: "0.1.0", scheduleTz: SCHEDULE_TZ, today: todayPT, nowMinutes: nowPT[0] * 60 + nowPT[1], source: "real"
  },
  master, systems, jobs, report
};
const banner = "/* Generated by scripts/build-mockup-data.mjs from " + path.basename(SRC) + " on " + now.toISOString() + ".\n * Contains real workbook data: do not commit. Re-run the script after the workbook changes. */\n";
fs.writeFileSync(OUT, banner + "window.CATALOG = " + JSON.stringify(data) + ";\n");

// Console report
const t = report.totals;
console.log("Wrote " + path.relative(ROOT, OUT) + " (" + Math.round(fs.statSync(OUT).size / 1024) + " KB)");
console.log("Tabs: " + report.tabs.map((x) => x.name + "=" + x.rows).join(", "));
console.log("Skipped: " + report.skipped.map((x) => x.name + " (" + x.reason + ")").join(", "));
console.log(`Jobs ${t.jobs} · categories ${t.categories} · services ${t.services} · fully matched ${t.fullyMatched} · suggested ${t.suggested} · inferred ${t.inferred} · unassigned ${t.unassigned} · unmapped ${t.unmapped}`);
console.log(`Frequencies: parsed ${t.parsed} · assumed ${t.assumed} · unparseable ${t.unparseable} · unscheduled ${t.unscheduled} · numeric cells ${t.numericCells}`);
if (process.argv.includes("--verbose")) {
  const seenF = new Map(); jobs.forEach((j) => { const k = (j.freq || "(blank)") + (j.cronCol ? "  [cron " + j.cronCol + "]" : ""); if (!seenF.has(k)) seenF.set(k, [j, 0]); seenF.get(k)[1]++; });
  [...seenF].sort((a, b) => a[1][0].conf.localeCompare(b[1][0].conf)).forEach(([k, [j, n]]) => console.log(String(n).padStart(3) + "  " + j.conf.padEnd(11) + j.rule.padEnd(16) + JSON.stringify(k).padEnd(70) + " → " + j.human));
  report.issues.forEach((i) => console.log(i.level.padEnd(8) + i.code.padEnd(30) + String(i.count).padStart(4) + "  e.g. " + i.examples.slice(0, 2).join(" | ")));
}
