/*
 * Shared mockup runtime. One helper per component in DESIGN.md §13.4, plus a tiny version of
 * the schedule engine (§11) so the calendar shows real per-day aggregates.
 * Data: window.CATALOG from catalog-data.js (generated from the real workbook by scripts/build-mockup-data.mjs).
 * Classic script (no modules), so it works on file://. All workbook text goes through esc().
 */
(function () {
  "use strict";
  var S = window.CATALOG;
  var RFV = {};
  window.RFV = RFV;

  /* ── Utilities ─────────────────────────────────────────── */
  var ESC = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" };
  function esc(v) { return v == null ? "" : String(v).replace(/[&<>"']/g, function (c) { return ESC[c]; }); }
  RFV.esc = esc;
  RFV.$ = function (sel, root) { return (root || document).querySelector(sel); };
  RFV.$$ = function (sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); };
  function store(key, val) {
    try { if (val === undefined) return window.localStorage.getItem(key); window.localStorage.setItem(key, val); } catch (e) { /* storage blocked: fine */ }
    return null;
  }
  RFV.store = store;
  RFV.qs = function (k) { try { return new URLSearchParams(window.location.search).get(k); } catch (e) { return null; } };

  /* ── Icons (inline SVG, lucide-style) ──────────────────── */
  var P = {
    search: '<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>',
    file: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6"/>',
    refresh: '<path d="M21 12a9 9 0 1 1-3-6.7L21 8"/><path d="M21 3v5h-5"/>',
    download: '<path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M5 21h14"/>',
    upload: '<path d="M12 21V9"/><path d="m7 14 5-5 5 5"/><path d="M5 3h14"/>',
    theme: '<circle cx="12" cy="12" r="9"/><path d="M12 3a9 9 0 0 1 0 18z" fill="currentColor"/>',
    tree: '<rect x="9" y="3" width="6" height="5" rx="1"/><rect x="3" y="16" width="6" height="5" rx="1"/><rect x="15" y="16" width="6" height="5" rx="1"/><path d="M12 8v4M6 16v-2h12v2"/>',
    jobs: '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><path d="M6.5 10v4a3 3 0 0 0 3 3H14"/>',
    calendar: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
    table: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M3 15h18M9 4v16"/>',
    warn: '<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
    check: '<path d="m5 12 5 5L20 7"/>',
    x: '<path d="M18 6 6 18M6 6l12 12"/>',
    chevDown: '<path d="m6 9 6 6 6-6"/>',
    chevL: '<path d="m15 18-6-6 6-6"/>',
    chevR: '<path d="m9 18 6-6-6-6"/>',
    clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    plus: '<path d="M12 5v14M5 12h14"/>',
    minus: '<path d="M5 12h14"/>',
    fit: '<path d="M8 3H5a2 2 0 0 0-2 2v3M21 8V5a2 2 0 0 0-2-2h-3M3 16v3a2 2 0 0 0 2 2h3M16 21h3a2 2 0 0 0 2-2v-3"/>',
    columns: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M9 4v16M15 4v16"/>',
    link: '<path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1"/><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/>',
    diamond: '<path d="M12 3 21 12 12 21 3 12z" fill="currentColor" stroke="none"/>',
    cluster: '<path d="M4 7h16M4 12h16M4 17h16"/>',
    grid: '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',
    expand: '<path d="m7 15 5 5 5-5M7 9l5-5 5 5"/>',
    collapse: '<path d="m7 20 5-5 5 5M7 4l5 5 5-5"/>',
    printer: '<path d="M6 9V3h12v6"/><rect x="3" y="9" width="18" height="8" rx="2"/><path d="M7 14h10v7H7z"/>',
    eye: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
    info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7h.01"/>'
  };
  function icon(name, cls) {
    return '<svg class="' + (cls || "i") + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + (P[name] || "") + "</svg>";
  }
  RFV.icon = icon;

  /* ── Severity (icon + label, never colour alone) ───────── */
  var SEV_ORDER = { Sev1: 0, Sev2: 1, Sev3: 2, Unknown: 3 };
  RFV.SEV_ORDER = SEV_ORDER;
  RFV.SEVS = ["Sev1", "Sev2", "Sev3", "Unknown"];
  var mark = 'style="stroke:var(--surface)" stroke-width="1.8" stroke-linecap="round"';
  var SEV_SVG = {
    Sev1: '<path d="M5.2 1h5.6L15 5.2v5.6L10.8 15H5.2L1 10.8V5.2z" fill="currentColor"/><path d="M8 4.4v4.3M8 11.3v.1" ' + mark + "/>",
    Sev2: '<path d="M8 1.2 15.2 14.5H.8z" fill="currentColor"/><path d="M8 6v3.6M8 12v.1" ' + mark + "/>",
    Sev3: '<circle cx="8" cy="8" r="7" fill="currentColor"/><path d="M8 7.3v4.4M8 4.6v.1" ' + mark + "/>",
    Unknown: '<circle cx="8" cy="8" r="6.2" fill="none" stroke="currentColor" stroke-width="1.6" stroke-dasharray="2.6 1.9"/><path d="M6.4 6.3a1.7 1.7 0 1 1 2.4 1.6c-.5.3-.8.6-.8 1.2M8 11.2v.1" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>'
  };
  RFV.sevIcon = function (sev, cls) {
    return '<svg class="' + (cls || "i-sm") + ' sev-c-' + sev + '" viewBox="0 0 16 16" aria-hidden="true">' + SEV_SVG[sev] + "</svg>";
  };
  RFV.sevBadge = function (sev) {
    sev = sev || "Unknown";
    return '<span class="sev sev-' + sev + '" title="Severity: ' + sev + '"><svg viewBox="0 0 16 16" aria-hidden="true">' + SEV_SVG[sev] + "</svg>" + sev + "</span>";
  };
  RFV.maxSev = function (jobs) {
    var best = null;
    jobs.forEach(function (j) { if (best == null || SEV_ORDER[j.sev] < SEV_ORDER[best]) best = j.sev; });
    return best;
  };

  /* ── SystemChip ────────────────────────────────────────── */
  RFV.sysChip = function (sysId, sub) {
    var s = S.systems[sysId] || { name: sysId, glyph: "?" };
    return '<span class="sys sys-' + sysId + '"><span class="glyph" aria-hidden="true">' + s.glyph + "</span>" + esc(s.name) +
      (sub ? ' <span class="subsys">· ' + esc(sub) + "</span>" : "") + "</span>";
  };

  /* ── StateTag / MonitoredFlag / ConfidenceMark / LinkStatusTag ─ */
  RFV.stateTag = function (st) {
    return st === "Enabled" ? '<span class="tag tag-ok">Enabled</span>' : st === "Disabled" ? '<span class="tag">Disabled</span>' : '<span class="tag">State unknown</span>';
  };
  RFV.monFlag = function (m, compact) {
    if (m === true) return '<span class="tag tag-ok" title="Monitored">' + icon("check", "i-sm") + (compact ? "Mon" : "Monitored") + "</span>";
    if (m === false) return '<span class="tag tag-warn" title="Not monitored">' + icon("warn", "i-sm") + "Unmonitored</span>";
    return '<span class="tag" title="Monitored column is blank">Monitored?</span>';
  };
  RFV.confMark = function (job) {
    if (job.conf === "assumed") return '<span class="conf conf-assumed" title="Assumed: ' + esc((job.assumptions || ["some details defaulted"]).join("; ")) + '">≈</span>';
    if (job.conf === "unparseable") return '<span class="conf conf-unparseable" title="Could not interpret this frequency">?</span>';
    return "";
  };
  RFV.freqLabel = function (job) {
    var raw = job.source === "cronColumn" ? job.freq + "  ·  cron column: " + job.cronCol : job.freq;
    return '<span class="freq" title="Raw: ' + esc(raw) + '">' + icon("clock", "i-sm") + '<span class="txt">' + esc(job.human) + "</span>" + RFV.confMark(job) + "</span>";
  };
  var LINK = { exact: ["tag-ok", "exact match"], normalized: ["tag-ok", "normalized match"], mapped: ["tag-info", "mapped"],
    inferred: ["tag-info", "inferred"], suggested: ["tag-dotted", "suggested mapping"], unmatched: ["tag-warn", "unmatched"] };
  RFV.linkTag = function (status) { var l = LINK[status]; return '<span class="tag ' + l[0] + '">' + l[1] + "</span>"; };

  /* ── Regions and placement ─────────────────────────────── */
  var ALL_REGIONS = ["US", "CA", "AU", "JP"];
  RFV.ALL_REGIONS = ALL_REGIONS;
  RFV.regionsOf = function (job) {
    if (!job.region) return [];
    var out = [];
    job.region.split(/[,&/;()]|\band\b/i).forEach(function (t) {
      t = t.trim().toUpperCase();
      if (!t) return;
      if (t === "GLOBAL" || t === "ALL") ALL_REGIONS.forEach(function (r) { if (out.indexOf(r) < 0) out.push(r); });
      else if (out.indexOf(t) < 0) out.push(t);
    });
    return out;
  };
  RFV.catName = function (j) { return j.place.cat || "Unmapped"; };
  RFV.svcName = function (j) {
    if (!j.place.cat) return j.catRaw ? 'Unknown category "' + j.catRaw.trim() + '"' : "(blank category)";
    return j.place.svc || "Unassigned service";
  };
  RFV.isUnplaced = function (j) { return !j.place.cat || !j.place.svc; };
  var JOB_BY_ID = {}; S.jobs.forEach(function (j) { JOB_BY_ID[j.id] = j; });
  RFV.job = function (id) { return JOB_BY_ID[id] || null; };

  /* ── JobTile ───────────────────────────────────────────── */
  RFV.jobTile = function (j) {
    var cls = "tile" + (j.state === "Disabled" ? " is-disabled" : "") + (j.status === "suggested" ? " is-suggested" : "");
    return '<button type="button" class="' + cls + '" data-job="' + esc(j.id) + '" aria-label="' + esc(j.name) + ', open details">' +
      '<span class="row1">' + RFV.sysChip(j.sys, j.sub) + RFV.sevBadge(j.sev) + "</span>" +
      '<span class="name" title="' + esc(j.name) + '">' + esc(j.name) + "</span>" +
      RFV.freqLabel(j) +
      '<span class="meta-line">' + esc(j.region || "No region") + '<span class="dot"></span>' + (j.state === "Enabled" ? "Enabled" : esc(j.state)) +
      '<span class="dot"></span>' + (j.mon === false ? '<span style="color:var(--warn)">' + icon("warn", "i-sm") + " Unmonitored</span>" : j.mon === true ? "Monitored" : "Monitored?") + "</span>" +
      "</button>";
  };

  /* ── Theme ─────────────────────────────────────────────── */
  function applyTheme(t) {
    if (t === "light" || t === "dark") document.documentElement.setAttribute("data-theme", t);
    else document.documentElement.removeAttribute("data-theme");
  }
  applyTheme(store("rfv-theme"));
  RFV.cycleTheme = function () {
    var cur = store("rfv-theme") || "system";
    var next = cur === "system" ? "light" : cur === "light" ? "dark" : "system";
    store("rfv-theme", next); applyTheme(next);
    var b = document.getElementById("theme-btn"); if (b) b.setAttribute("title", "Theme: " + next);
    var l = document.getElementById("theme-lbl"); if (l) l.textContent = next[0].toUpperCase() + next.slice(1);
  };

  /* ── Display timezone ──────────────────────────────────── */
  RFV.TZS = [["America/Los_Angeles", "Pacific (LA)"], ["America/New_York", "Eastern (NY)"], ["America/Toronto", "Toronto"], ["Australia/Sydney", "Sydney"], ["Asia/Tokyo", "Tokyo"], ["UTC", "UTC"]];
  RFV.tz = function () { return store("rfv-tz") || "America/Los_Angeles"; };
  RFV.tzLabel = function (tz) { for (var i = 0; i < RFV.TZS.length; i++) if (RFV.TZS[i][0] === tz) return RFV.TZS[i][1]; return tz; };

  /* ── AppShell ──────────────────────────────────────────── */
  var NAV = [
    ["load-workbook.html", "upload", "Load workbook"],
    ["services-tree.html", "tree", "Business Services"],
    ["service-jobs-tree.html", "jobs", "Service → Jobs"],
    ["calendar.html", "calendar", "Job calendar"],
    ["jobs-explorer.html", "table", "Jobs Explorer"],
    ["gaps.html", "warn", "Gaps & Data Quality", String(S.jobs.filter(function (j) { return !j.place.cat || !j.place.svc || j.status === "suggested"; }).length)]
  ];
  RFV.shell = function (opts) {
    var navHtml = NAV.map(function (n) {
      return '<a href="' + n[0] + '"' + (n[0] === opts.page ? ' aria-current="page"' : "") + ' title="' + n[2] + '">' + icon(n[1]) + '<span class="lbl">' + n[2] + "</span>" + (n[3] ? '<span class="count">' + n[3] + "</span>" : "") + "</a>";
    }).join("");
    var theme = store("rfv-theme") || "system";
    var tzOpts = RFV.TZS.map(function (t) { return '<option value="' + t[0] + '"' + (t[0] === RFV.tz() ? " selected" : "") + ">" + t[1] + "</option>"; }).join("");
    document.body.innerHTML =
      '<a class="sr-only" href="#main">Skip to content</a>' +
      '<div class="app">' +
      '<header class="topbar">' +
      '<div class="brand"><span class="brand-mark">' + icon("grid", "i-sm") + "</span>Service–Job Catalog</div>" +
      '<div class="wb-pill" title="' + esc(S.meta.workbookPath) + '">' + icon("file", "i-sm") + "<strong>" + esc(S.meta.workbook) + '</strong><span class="mod">· modified ' + esc(S.meta.modified) + "</span>" +
      '<button class="btn btn-ghost btn-sm btn-icon" title="Reload workbook" aria-label="Reload workbook">' + icon("refresh", "i-sm") + "</button></div>" +
      '<span class="spacer"></span>' +
      '<label class="check" title="Display timezone">' + icon("clock", "i-sm") + '<span class="sr-only">Display timezone</span><select class="select" id="tz-select" style="height:30px">' + tzOpts + "</select></label>" +
      '<button class="btn btn-icon" id="theme-btn" title="Theme: ' + theme + '" aria-label="Toggle theme">' + icon("theme") + "</button>" +
      (opts.exportView ? '<a class="btn btn-primary" href="export-dialog.html?view=' + opts.exportView + '">' + icon("download", "i-sm") + "Export as HTML</a>" : "") +
      "</header>" +
      '<nav class="nav" aria-label="Views">' + navHtml + '<div class="nav-sep"></div><a href="index.html" title="Mockup index">' + icon("info") + '<span class="lbl">Mockup index</span></a><p class="nav-note">Mockup · real workbook data</p></nav>' +
      '<main class="main" id="main" tabindex="-1">' +
      (opts.title ? '<div class="page-head"><h1>' + esc(opts.title) + "</h1>" + (opts.sub ? '<span class="sub">' + opts.sub + "</span>" : "") + '<span style="flex:1"></span>' + (opts.headRight || "") + "</div>" : "") +
      '<div id="view" style="display:flex;flex-direction:column;flex:1 0 auto"></div>' +
      '<div class="mock-banner">Mockup for design and data review · data from ' + esc(S.meta.workbook) + " (modified " + esc(S.meta.modified) + ") · " + S.jobs.length + " jobs, " + S.master.length + " categories, " + S.report.totals.services + " services · generated " + esc(S.meta.loadedAt) + ' · matches DESIGN.md §13–§15</div>' +
      "</main></div>" +
      '<div class="overlay" id="overlay"></div>' +
      '<aside class="drawer" id="drawer" role="dialog" aria-modal="true" aria-labelledby="drawer-title" aria-hidden="true"></aside>';
    document.getElementById("theme-btn").addEventListener("click", RFV.cycleTheme);
    document.getElementById("tz-select").addEventListener("change", function (e) {
      store("rfv-tz", e.target.value);
      document.dispatchEvent(new CustomEvent("rfv:tz", { detail: e.target.value }));
    });
    document.getElementById("overlay").addEventListener("click", RFV.closeDrawer);
    document.addEventListener("keydown", function (e) { if (e.key === "Escape" && RFV.closeDrawer()) e.rfvDrawerClosed = true; });
    document.addEventListener("click", function (e) {
      var t = e.target.closest("[data-job]");
      if (t) { RFV.openJob(t.getAttribute("data-job")); return; }
      RFV.$$("details.fchip[open]").forEach(function (d) { if (!d.contains(e.target)) d.removeAttribute("open"); });
    });
    return document.getElementById("view");
  };

  /* ── DetailDrawer ──────────────────────────────────────── */
  var lastFocus = null;
  function kv(rows) {
    return '<dl class="kv">' + rows.map(function (r) { return "<dt>" + r[0] + "</dt><dd>" + (r[1] == null || r[1] === "" ? '<span class="unset">not set</span>' : r[1]) + "</dd>"; }).join("") + "</dl>";
  }
  RFV.openPanel = function (titleHtml, chipsHtml, bodyHtml) {
    var d = document.getElementById("drawer");
    lastFocus = document.activeElement;
    d.innerHTML = '<header><h2 id="drawer-title"><span>' + titleHtml + '</span><button class="btn btn-ghost btn-icon" id="drawer-close" aria-label="Close details">' + icon("x") + "</button></h2>" +
      (chipsHtml ? '<div class="chips">' + chipsHtml + "</div>" : "") + '</header><div class="body">' + bodyHtml + "</div>";
    d.classList.add("open"); d.setAttribute("aria-hidden", "false");
    document.getElementById("overlay").classList.add("open");
    document.getElementById("drawer-close").addEventListener("click", RFV.closeDrawer);
    document.getElementById("drawer-close").focus();
  };
  RFV.closeDrawer = function () {
    var d = document.getElementById("drawer");
    if (!d || !d.classList.contains("open")) return false;
    d.classList.remove("open"); d.setAttribute("aria-hidden", "true");
    document.getElementById("overlay").classList.remove("open");
    if (lastFocus && lastFocus.focus) lastFocus.focus();
    return true;
  };
  RFV.openJob = function (id) {
    var j = RFV.job(id); if (!j) return;
    var tz = RFV.tz();
    var next = RFV.nextRuns(j, tz, 4);
    var alertKind = function (a) { var m = /_(Failed|Not_?Finished|Not_?Started|Error|Warning|NotCompleted|Failure)$/i.exec(a); return m ? m[1].replace("_", "") : ""; };
    var alerts = j.alerts.length
      ? '<ul class="alert-list">' + j.alerts.map(function (a) { var k = alertKind(a); return "<li><span>" + esc(a) + "</span>" + (k ? '<span class="tag">' + esc(k) + "</span>" : "") + "</li>"; }).join("") + "</ul>"
      : '<span class="unset">no alerts</span>' + (j.alertsRaw ? ' · sheet says <span class="raw">' + esc(j.alertsRaw) + "</span>" : "");
    var dep = j.sched.t === "dep" ? RFV.job(j.sched.up) : null;
    var attrs = j.attrs ? Object.keys(j.attrs).map(function (k) { return [esc(k), j.attrs[k] ? esc(j.attrs[k]) : '<span class="unset">blank</span>']; }) : [];
    var body =
      '<section class="dsec"><h3>Overview</h3>' + kv([
        ["Description", esc(j.desc)], ["System", RFV.sysChip(j.sys, j.sub)], ["Region", esc(j.region) + (j.region ? ' <span class="faint">→ ' + RFV.regionsOf(j).join(", ") + "</span>" : "")],
        ["State", RFV.stateTag(j.state) + (j.stateRaw && j.stateRaw !== j.state ? ' <span class="raw">' + esc(j.stateRaw) + "</span>" : "")],
        ["Monitored", RFV.monFlag(j.mon) + (j.monRaw ? ' <span class="faint">sheet: ' + esc(j.monRaw) + "</span>" : "")]]) + "</section>" +
      '<section class="dsec"><h3>Schedule</h3>' + kv([
        ["Raw text", j.freq ? '<span class="raw">' + esc(j.freq) + "</span>" : '<span class="unset">blank</span>'],
        j.cronCol ? ["Text reads as", esc(j.textHuman || "") + ' <span class="faint">(not used)</span>'] : null,
        j.cronCol ? ["Cron column", '<span class="raw">' + esc(j.cronCol) + '</span> <span class="faint">(Schedule interval, used)</span>'] : null,
        j.conflict ? ["Conflict", '<span class="tag tag-warn">' + icon("warn", "i-sm") + esc(j.conflict) + "</span>"] : null,
        ["Interpreted", esc(j.human) + ' <span class="faint">(Pacific)</span>'],
        ["Confidence", (j.conf === "parsed" ? '<span class="tag tag-ok">parsed</span>' : j.conf === "assumed" ? '<span class="tag tag-warn">assumed</span>' : '<span class="tag tag-danger">unparseable</span>') + ' <span class="faint mono">rule: ' + esc(j.rule) + "</span>"],
        j.assumptions ? ["Assumptions", esc(j.assumptions.join("; "))] : null,
        dep ? ["Upstream", '<a href="#" data-job="' + dep.id + '">' + esc(dep.name) + "</a>"] : null,
        ["Next runs", next.length ? next.map(esc).join(", ") + ' <span class="faint">(' + esc(RFV.tzLabel(tz)) + ")</span>" : (RFV.isScheduled(j) ? '<span class="unset">none in the next 40 days</span>' : '<span class="unset">none: unscheduled</span>')]
      ].filter(Boolean)) + "</section>" +
      '<section class="dsec"><h3>Service link</h3>' + kv([
        ["Category (sheet)", j.catRaw ? '<span class="raw">' + esc(j.catRaw) + "</span>" : '<span class="unset">blank</span>'],
        ["Service (sheet)", j.svcRaw ? '<span class="raw">' + esc(j.svcRaw) + "</span>" : '<span class="unset">blank</span>'],
        ["Placed under", esc(RFV.catName(j)) + " / " + esc(RFV.svcName(j))],
        ["Status", RFV.linkTag(j.status) + (j.status === "suggested" ? ' <a href="gaps.html#mapping">confirm on Gaps page</a>' : "")]]) + "</section>" +
      '<section class="dsec"><h3>Alerts (' + j.alerts.length + ")</h3>" + alerts + '<div style="margin-top:8px">' + kv([["Alert frequency", esc(j.alertFreq)]]) + "</div></section>" +
      '<section class="dsec"><h3>Ownership</h3>' + kv([
        ["Severity", j.sev === "Unknown" ? RFV.sevBadge("Unknown") + (j.sevRaw ? ' <span class="raw">' + esc(j.sevRaw) + "</span>" : ' <span class="unset">not set in sheet</span>') : RFV.sevBadge(j.sev) + ' <span class="faint">sheet: ' + esc(j.sevRaw) + "</span>"],
        ["Resolution SLA", esc(j.sla)], ["L1 responder", esc(j.l1)], ["L2 resolver", esc(j.l2)], ["R+F stakeholder", esc(j.owner)]]) + "</section>" +
      '<section class="dsec"><h3>Impact</h3>' + kv([["System impact", esc(j.sysImpact)], ["Business impact", esc(j.bizImpact)], ["Impacted user group", esc(j.group)]]) + "</section>" +
      (attrs.length ? '<section class="dsec"><h3>Additional attributes</h3>' + kv(attrs) + "</section>" : "") +
      '<section class="dsec"><h3>Provenance</h3>' + kv([["Sheet", esc(j.tab)], ["Row", String(j.row)], ["Job id", '<span class="mono">' + esc(j.id) + "</span>"]]) + "</section>";
    var chips = RFV.sysChip(j.sys, j.sub) + RFV.sevBadge(j.sev) + RFV.stateTag(j.state) + RFV.monFlag(j.mon) + '<span class="tag">' + esc(j.region || "No region") + "</span>";
    RFV.openPanel(esc(j.name), chips, body);
  };

  /* ── FilterBar ─────────────────────────────────────────── */
  RFV.filterChip = function (cfg) {
    var d = document.createElement("details");
    d.className = "fchip";
    function render() {
      var n = cfg.selected.size;
      d.setAttribute("data-active", n ? "true" : "false");
      var lbl = n === 0 ? cfg.label : n === 1 ? cfg.label + ": " + labelOf(Array.from(cfg.selected)[0]) : cfg.label + ": " + n;
      d.innerHTML = "<summary>" + esc(lbl) + icon("chevDown", "i-sm") + '</summary><div class="menu" role="group" aria-label="' + esc(cfg.label) + '">' +
        cfg.options.map(function (o) { return '<label><input type="checkbox" value="' + esc(o.v) + '"' + (cfg.selected.has(o.v) ? " checked" : "") + ">" + (o.html || esc(o.l)) + "</label>"; }).join("") + "</div>";
    }
    function labelOf(v) { for (var i = 0; i < cfg.options.length; i++) if (cfg.options[i].v === v) return cfg.options[i].l; return v; }
    d.addEventListener("change", function (e) {
      var v = e.target.value;
      if (e.target.checked) cfg.selected.add(v); else cfg.selected.delete(v);
      var open = d.open; render(); d.open = open; cfg.onChange();
    });
    render();
    return d;
  };
  function distinct(fn) {
    var seen = {}, out = [];
    S.jobs.forEach(function (j) { var v = fn(j); if (!seen[v]) { seen[v] = 1; out.push({ v: v, l: v }); } });
    return out.sort(function (a, b) { return a.l.localeCompare(b.l); });
  }
  RFV.distinct = distinct;
  RFV.filterOptions = {
    sys: Object.keys(S.systems).map(function (k) { return { v: k, l: S.systems[k].name, html: RFV.sysChip(k) }; }),
    sev: ["Sev1", "Sev2", "Sev3", "Unknown"].map(function (s) { return { v: s, l: s, html: RFV.sevBadge(s) }; }),
    state: [{ v: "Enabled", l: "Enabled" }, { v: "Disabled", l: "Disabled" }],
    mon: [{ v: "yes", l: "Monitored" }, { v: "no", l: "Unmonitored" }, { v: "unknown", l: "Blank" }],
    region: ALL_REGIONS.map(function (r) { return { v: r, l: r }; }),
    cat: S.master.map(function (c) { return { v: c[0], l: c[0] }; }).concat([{ v: "Unmapped", l: "Unmapped" }]),
    l1: distinct(function (j) { return j.l1 || "(blank)"; }),
    l2: distinct(function (j) { return j.l2 || "(blank)"; })
  };
  RFV.passes = function (j, f) {
    if (f.sys && f.sys.size && !f.sys.has(j.sys)) return false;
    if (f.sev && f.sev.size && !f.sev.has(j.sev)) return false;
    if (f.state && f.state.size && !f.state.has(j.state)) return false;
    if (f.mon && f.mon.size && !f.mon.has(j.mon === true ? "yes" : j.mon === false ? "no" : "unknown")) return false;
    if (f.region && f.region.size && !RFV.regionsOf(j).some(function (r) { return f.region.has(r); })) return false;
    if (f.cat && f.cat.size && !f.cat.has(RFV.catName(j))) return false;
    if (f.l1 && f.l1.size && !f.l1.has(j.l1 || "(blank)")) return false;
    if (f.l2 && f.l2.size && !f.l2.has(j.l2 || "(blank)")) return false;
    return true;
  };
  RFV.filterText = function (f, names) {
    var parts = [];
    Object.keys(names).forEach(function (k) {
      if (f[k] && f[k].size) parts.push(names[k] + ": " + Array.from(f[k]).map(function (v) { return k === "sys" ? S.systems[v].name : v; }).join(", "));
    });
    return parts.length ? parts.join(" · ") : "No filters";
  };

  /* ── Schedule engine (mini) ────────────────────────────── */
  // Dates are UTC-midnight Date objects used as calendar days.
  RFV.day = function (y, m, d) { return new Date(Date.UTC(y, m - 1, d)); };
  RFV.addDays = function (dt, n) { return new Date(dt.getTime() + n * 86400000); };
  RFV.key = function (dt) { return dt.toISOString().slice(0, 10); };
  function nthOk(dt, nth) {
    if (dt.getUTCDay() !== nth.wd) return false;
    return Math.ceil(dt.getUTCDate() / 7) === nth.n;
  }
  function matchesDay(s, dt) {
    if (s.dom && s.dom.indexOf(dt.getUTCDate()) < 0) return false;
    if (s.dow && s.dow.indexOf(dt.getUTCDay()) < 0) return false;
    if (s.months && s.months.indexOf(dt.getUTCMonth() + 1) < 0) return false;
    if (s.nth && !nthOk(dt, s.nth)) return false;
    if (s.dstep) { var anchor = Date.UTC(2026, 0, 1); if (Math.round((dt.getTime() - anchor) / 86400000) % s.dstep !== 0) return false; }
    return true;
  }
  function resolve(job, depth) {
    if (job.sched.t !== "dep") return job.sched;
    if ((depth || 0) > 5) return { t: "none" };
    var up = RFV.job(job.sched.up); return up ? resolve(up, (depth || 0) + 1) : { t: "none" };
  }
  RFV.isScheduled = function (j) { return resolve(j).t !== "none"; };
  RFV.isDaily = function (j) { var s = resolve(j); return s.t === "int" || (s.t === "cal" && !s.dom && !s.dow && !s.months && !s.nth && !s.dstep); };
  function runsSource(job, dt) {
    var s = resolve(job), out = [];
    if (s.t === "int") { for (var sec = 0; sec < 86400; sec += s.every) out.push(Math.floor(sec / 60)); return out; }
    if (s.t !== "cal" || !matchesDay(s, dt)) return out;
    for (var a = 0; a < s.h.length; a++) for (var b = 0; b < s.m.length; b++) out.push(s.h[a] * 60 + s.m[b]);
    return out.sort(function (x, y) { return x - y; });
  }
  var fmtCache = {};
  function offsetMin(tz, ms) {
    var f = fmtCache[tz] || (fmtCache[tz] = new Intl.DateTimeFormat("en-US", { timeZone: tz, hourCycle: "h23", year: "numeric", month: "numeric", day: "numeric", hour: "numeric", minute: "numeric" }));
    var p = {}; f.formatToParts(new Date(ms)).forEach(function (x) { p[x.type] = x.value; });
    return (Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour % 24, +p.minute) - ms) / 60000;
  }
  function delta(dt, tz) { var ms = dt.getTime() + 20 * 3600000; return offsetMin(tz, ms) - offsetMin(S.meta.scheduleTz, ms); }
  /* Aggregate runs per display day: Map<dateKey, Map<jobId, {count, first, hist[24]}>> (DESIGN.md §11) */
  RFV.aggregate = function (jobs, from, to, tz, keepTimes) {
    var res = {};
    for (var dt = RFV.addDays(from, -1); dt <= RFV.addDays(to, 1); dt = RFV.addDays(dt, 1)) {
      var dl = delta(dt, tz);
      jobs.forEach(function (j) {
        var runs = runsSource(j, dt);
        for (var i = 0; i < runs.length; i++) {
          var t = runs[i] + dl, off = Math.floor(t / 1440), mm = t - off * 1440;
          var k = RFV.key(RFV.addDays(dt, off));
          var day = res[k] || (res[k] = {});
          var a = day[j.id] || (day[j.id] = { count: 0, first: mm, last: mm, hist: new Array(24).fill(0), times: [] });
          a.count++; if (mm < a.first) a.first = mm; if (mm > a.last) a.last = mm; a.hist[Math.floor(mm / 60)]++;
          if (a.times) { a.times.push(mm); if (!keepTimes && a.times.length > 48) a.times = null; }
        }
      });
    }
    return res;
  };
  RFV.hhmm = function (m) { return String(Math.floor(m / 60)).padStart(2, "0") + ":" + String(m % 60).padStart(2, "0"); };
  RFV.todayKey = function () { return S.meta.today; };
  RFV.nextRuns = function (j, tz, n) {
    var t = S.meta.today.split("-").map(Number), d0 = RFV.day(t[0], t[1], t[2]);
    var agg = RFV.aggregate([j], d0, RFV.addDays(d0, 40), tz, true), out = [];
    for (var i = 0; i <= 40 && out.length < n; i++) {
      var k = RFV.key(RFV.addDays(d0, i)), a = agg[k] && agg[k][j.id];
      if (!a) continue;
      a.times.sort(function (x, y) { return x - y; });
      for (var q = 0; q < a.times.length && out.length < n; q++) {
        if (i === 0 && a.times[q] < S.meta.nowMinutes) continue; // runs still ahead today
        out.push((i === 0 ? "" : k.slice(5) + " ") + RFV.hhmm(a.times[q]));
      }
    }
    return out;
  };
})();
