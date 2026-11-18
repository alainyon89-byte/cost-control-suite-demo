// (c) 2026 Alain Yon
"use strict";

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const TOKEN = (location.hash.match(/t=([^&]+)/) || [])[1] || "";

const S = {
  st: null, mode: { mode: "window", drop: true }, route: { name: "home" }, search: "",
  tool: null, files: {}, opts: {}, shared: {}, checks: [], manual: {},
  run: null, lastTool: null, hist: { tool: "", month: "" }, setup: null, sap: null,
};

// ---------- talking to Python ----------
function pyReady() {
  return new Promise((resolve) => {
    if (TOKEN) return resolve();
    if (window.pywebview && window.pywebview.api) return resolve();
    window.addEventListener("pywebviewready", () => resolve(), { once: true });
  });
}
async function api(name, ...args) {
  if (!TOKEN) return await window.pywebview.api.call(name, args);
  const r = await fetch("/api/call", { method: "POST", headers: { "Content-Type": "application/json", "X-CCS-Token": TOKEN }, body: JSON.stringify({ name, args }) });
  return await r.json();
}

// ---------- small helpers ----------
function toast(msg, ms = 4500) {
  const el = document.createElement("div");
  el.className = "toast"; el.textContent = msg;
  $("#toast").appendChild(el);
  setTimeout(() => el.remove(), ms);
}
function modal(html) { const m = $("#modal"); m.innerHTML = `<div class="dialog" role="dialog" aria-modal="true">${html}</div>`; m.hidden = false; const f = $("button, input", m); if (f) f.focus(); }
function closeModal() { $("#modal").hidden = true; $("#modal").innerHTML = ""; }
async function copyText(text) {
  try { await navigator.clipboard.writeText(text); }
  catch (e) { const t = document.createElement("textarea"); t.value = text; document.body.appendChild(t); t.select(); document.execCommand("copy"); t.remove(); }
  toast("Copied");
}
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function niceDate(iso, withTime = true) {
  if (!iso) return "";
  const m = String(iso).match(/^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d))?/);
  if (!m) return String(iso);
  return `${+m[3]} ${MONTHS[+m[2] - 1]} ${m[1]}` + (withTime && m[4] ? `, ${m[4]}:${m[5]}` : "");
}
const STATUS_WORD = { finished: "Finished", failed: "Failed", stopped: "Stopped", cancelled: "Cancelled", approval_gate: "Needs approval" };
function statusPill(status) {
  const cls = status === "finished" ? "ok" : status === "failed" ? "bad" : "warn";
  return `<span class="pill ${cls}">${status === "finished" ? "✓ " : ""}${esc(STATUS_WORD[status] || status)}</span>`;
}
function toolById(id) { return (S.st.tools || []).find((t) => t.id === id); }
function groupName(id) { return (S.st.groups.find((g) => g.id === id) || {}).name || id; }
function ageText(info) {
  if (!info || !info.loaded) return "not loaded yet";
  const d = info.age_days;
  return `exported ${niceDate(info.exported, false)} · ${d === 0 ? "today" : d === 1 ? "1 day old" : d + " days old"}`;
}

// ---------- navigation ----------
function go(route) { S.route = route; render(); $("#main").scrollTop = 0; $("#main").focus(); }
async function refresh() { S.st = await api("state"); }

function renderSide() {
  const r = S.route, cur = (name, extra) => (r.name === name && (!extra || r.group === extra) ? ' aria-current="page"' : "");
  const groups = S.st.groups.map((g) => {
    const n = S.st.tools.filter((t) => t.group === g.id).length;
    const active = (r.name === "group" && r.group === g.id) || (r.name === "tool" && S.tool && S.tool.group === g.id);
    return `<button data-act="nav-group" data-id="${g.id}"${active ? ' aria-current="page"' : ""}>${esc(g.name)}<span class="count">${n}</span></button>`;
  }).join("");
  $("#side").innerHTML = `
    <div class="brand">Cost Control Suite<span>${esc(S.st.app.property_name || "")}</span></div>
    <div class="nav">
      <button data-act="nav" data-to="home"${cur("home")}>Home</button>
      <button data-act="nav" data-to="sap"${cur("sap")}>SAP data</button>
      <div class="sep"></div>
      ${groups}
      <div class="sep"></div>
      <button data-act="nav" data-to="history"${cur("history")}>Run history</button>
      <button data-act="nav" data-to="settings"${cur("settings")}>Settings &amp; data folders</button>
    </div>
    <div class="side-foot">PR Tool runs as its own app<br>Version ${esc(S.st.app.version)}<br>© 2026 Alain Yon</div>`;
}

function render() {
  renderSide();
  const r = S.route;
  const main = $("#main");
  if (r.name === "home" || r.name === "group") main.innerHTML = viewHome();
  else if (r.name === "tool") main.innerHTML = viewTool();
  else if (r.name === "run") { main.innerHTML = viewRun(); afterRunRender(); }
  else if (r.name === "sap") main.innerHTML = viewSap();
  else if (r.name === "history") main.innerHTML = viewHistory();
  else if (r.name === "settings") main.innerHTML = viewSettings();
}

// ---------- Home ----------
function tile(t) {
  const last = t.last_run ? `${niceDate(t.last_run.time)} · ${STATUS_WORD[t.last_run.status] || t.last_run.status}` : "Not run yet";
  if (!t.installed) {
    return `<article class="tile off"><div class="tile-head"><h3>${esc(t.name)}</h3>${t.version ? `<span class="chip">${esc(t.version)}</span>` : ""}</div>
      <p class="what">${esc(t.what)}</p><p class="note warn"><strong>Not installed yet</strong> · add its folder and run the installer again</p></article>`;
  }
  return `<article class="tile"><div class="tile-head"><h3>${esc(t.name)}</h3>${t.version ? `<span class="chip">${esc(t.version)}</span>` : ""}</div>
    <p class="what">${esc(t.what)}</p>
    <dl><dt>Inputs</dt><dd>${esc(t.inputs_line)}</dd><dt>Last run</dt><dd>${esc(last)}</dd></dl>
    <button class="btn primary" data-act="open-tool" data-id="${t.id}">${t.running ? "Running: show" : "Open"}<span class="hide"> ${esc(t.name)}</span></button></article>`;
}
const SAP_ORDER = ["mm60", "mm60_cc", "mi20_kitchen"];
// plant codes and names come from config/app_settings.json (through the app state)
const PL = () => S.st.app.plants || { hotel: "D100", cc: "D200", names: {} };
const plantLabel = (code) => `${PL().names[code] || code} (${code})`;
const SAP_SHORT = new Proxy({}, { get: (_, k) => k === "mm60" ? "MM60 " + plantLabel(PL().hotel) : k === "mm60_cc" ? "MM60 " + plantLabel(PL().cc) : "Kitchen MI20" });
function slotPlant(spec) {
  const from = spec.plant_from || "";
  if (from.startsWith("opt:")) return S.opts[from.slice(4)] || PL().hotel;
  if (from.startsWith("file:")) { const f = (S.files[from.slice(5)] || [])[0]; return (f && f.plant) || PL().hotel; }
  return PL().hotel;
}
function sapLines() {
  return SAP_ORDER.map((k) => {
    const i = S.st.sap[k];
    const pill = !i.loaded ? '<span class="pill warn">Not loaded</span>' : i.old ? '<span class="pill warn">Getting old</span>' : '<span class="pill ok">✓ Current</span>';
    return `<div class="sapline"><strong>${esc(SAP_SHORT[k])}</strong>${pill}<span class="muted">${esc(ageText(i))}</span></div>`;
  }).join("");
}
function viewHome() {
  const r = S.route, q = S.search.trim().toLowerCase();
  const hour = new Date().getHours();
  const hello = hour < 12 ? "Good morning" : hour < 17 ? "Good afternoon" : "Good evening";
  const groups = S.st.groups.filter((g) => r.name !== "group" || g.id === r.group);
  let body = "";
  for (const g of groups) {
    const tools = S.st.tools.filter((t) => t.group === g.id && (!q || (t.name + " " + t.what + " " + t.inputs_line).toLowerCase().includes(q)));
    if (!tools.length) continue;
    body += `<div class="group-title"><h2>${esc(g.name)}</h2><span class="muted small">${tools.length} tool${tools.length === 1 ? "" : "s"}</span></div><div class="tiles">${tools.map(tile).join("")}</div>`;
  }
  if (!body) body = `<p class="muted" style="margin-top:24px">No tool matches “${esc(S.search)}”. <button class="link" data-act="clear-search">Show all tools</button></p>`;
  const top = r.name === "group" ? `<div class="hero"><div class="grow"><h1>${esc(groupName(r.group))}</h1></div></div>` : `
    <div class="hero"><div class="grow"><h1>${hello}, ${esc(S.st.app.first_name)}</h1><p>${esc(S.st.app.month)} · drop files anywhere on this page and the app suggests the tool.</p></div>
      <label class="search"><span class="hide">Find a tool</span><input type="search" id="search" placeholder="Find a tool" value="${esc(S.search)}"></label></div>
    <div class="home-top">
      <section class="card"><div class="row"><h2 class="grow">Month-end checklist · ${esc(S.st.app.month)}</h2><button class="link small" data-act="nav" data-to="settings">Edit</button></div>
        <ul class="checklist">${S.st.checklist.map((c) => `<li><span class="tick ${c.done ? "done" : ""}" aria-hidden="true">${c.done ? "✓" : ""}</span><button class="link" style="min-height:30px;text-decoration:none;color:inherit" data-act="open-tool" data-id="${c.id}">${esc(c.name)}</button><span class="muted small">${c.done ? "done" : "to do"}</span></li>`).join("")}</ul></section>
      <section class="card"><div class="row"><h2 class="grow">SAP data</h2><button class="btn" data-act="nav" data-to="sap">Open</button></div>${sapLines()}
        <p class="muted small" style="margin-top:6px">Loaded once here and used by every tool that needs it.</p></section>
    </div>`;
  return `<div class="wrap">${top}${body}</div>`;
}

// ---------- Tool page ----------
async function openTool(id, presetFiles) {
  const t = toolById(id);
  if (t && t.running && S.run && S.run.tool === id) return go({ name: "run" });
  S.tool = await api("tool_view", id);
  if (S.lastTool !== id || presetFiles) { S.files = {}; S.opts = {}; S.shared = {}; S.manual = {}; }
  S.lastTool = id;
  for (const o of S.tool.options) if (S.opts[o.key] === undefined) S.opts[o.key] = o.type === "checkbox" ? false : (o.default ?? "");
  if (S.tool.variant_option && S.opts[S.tool.variant_option] === undefined) {
    const first = S.tool.variants.find((v) => v.installed) || S.tool.variants[0];
    S.opts[S.tool.variant_option] = first.value;
  }
  S.checks = [];
  go({ name: "tool" });
  if (presetFiles && presetFiles.length) await addFiles(firstFileKey(), presetFiles);
  else runChecks();
}
function firstFileKey() { const s = S.tool.inputs.find((i) => !i.shared); return s ? s.key : null; }
function inputsPayload() {
  const out = {};
  for (const spec of S.tool.inputs) {
    if (spec.shared) {
      const sh = S.shared[spec.key];
      if (sh && sh.path) out[spec.key] = [sh.path];
      else if (sh && sh.mode) out[spec.key] = ["@" + sh.mode];
    } else out[spec.key] = (S.files[spec.key] || []).map((f) => f.path);
  }
  return out;
}
async function addFiles(key, paths) {
  const spec = S.tool.inputs.find((i) => i.key === key);
  if (!spec || !paths.length) return;
  const sniffed = await api("sniff", S.tool.id, paths);
  const cur = spec.multiple ? (S.files[key] || []) : [];
  for (const f of sniffed) if (!cur.some((c) => c.path === f.path)) cur.push(f);
  S.files[key] = spec.multiple ? cur : cur.slice(-1);
  render(); runChecks();
}
let checkTimer = null;
function runChecks() {
  clearTimeout(checkTimer);
  checkTimer = setTimeout(async () => {
    if (S.route.name !== "tool") return;
    S.checks = await api("preflight", S.tool.id, inputsPayload(), S.opts);
    const box = $("#checks"); if (box) box.innerHTML = checksHtml();
  }, 150);
}
function checksHtml() {
  if (!Array.isArray(S.checks)) return `<p class="note bad">${esc((S.checks && S.checks.error) || "The checks could not run.")}</p>`;
  const word = { ok: "✓ OK", fail: "✕ Fix this", warn: "! Check", manual: "You check" };
  const blocked = S.checks.some((c) => c.status === "fail");
  const rows = S.checks.map((c) => `<li><span class="mark ${c.status}">${word[c.status]}</span><div><div>${esc(c.text)}</div>${c.message ? `<div class="small ${c.status === "fail" ? "" : "muted"}" style="${c.status === "fail" ? "color:var(--bad)" : ""}">${esc(c.message)}</div>` : ""}</div></li>`).join("");
  return `<ul class="checks">${rows || '<li><span class="mark ok">✓ OK</span><div>Nothing to check for this tool.</div></li>'}</ul>
    <div class="row"><button class="btn primary big" data-act="run"${blocked ? " disabled" : ""}>${esc(S.tool.run_label || "Run")}</button>
    <button class="btn" data-act="recheck">Check again</button>${blocked ? '<span class="muted">Fix the items marked “Fix this” first.</span>' : ""}</div>`;
}
function sharedSlot(spec) {
  const plant = slotPlant(spec), sh = S.shared[spec.key] || {};
  const info = (spec.by_plant ? spec.by_plant[plant] : spec.shared_info) || {};
  const plantName = plantLabel(plant);
  const name = spec.shared === "mm60_previous" ? "Previous MM60 " + plantName : spec.shared === "mm60" ? "MM60 " + plantName : "Kitchen MI20";
  if (sh.path) {
    return `<div class="slot"><div class="grow"><strong>${esc(spec.label)}</strong><div>${esc(sh.name)} <span class="pill warn">This run only</span></div><div class="small muted">The shared file on the SAP data page is not changed.</div></div>
      <button class="btn" data-act="shared-reset" data-key="${spec.key}">Back to the shared file</button></div>`;
  }
  if (spec.deliver === "embedded_sheet") {
    const file = (S.files[S.tool.inputs.find((i) => !i.shared).key] || [])[0];
    const inFile = file && /has an MM60 sheet/.test(file.detail || "");
    const useShared = !!file && info.loaded && sh.mode !== "in_file";
    const sharedTxt = `MM60 ${plantName}: ${esc(info.original_name)} · ${esc(ageText(info))}`;
    return `<div class="slot"><div class="grow"><strong>${esc(spec.label)}</strong>
      <div>${!file ? "Add the Sales Mix first. Its plant decides which MM60 is used." : useShared ? `From SAP data · ${sharedTxt}` : inFile ? "Using the MM60 sheet inside your file" : "Your file has no MM60 sheet"}</div>
      <div class="small muted">${!file ? "" : useShared ? "It is written into a temporary copy of your file. Your file is not changed." : info.loaded ? "" : `The MM60 for ${plantName} is not loaded on the SAP data page.`}</div></div>
      ${file && inFile && info.loaded ? (useShared ? `<button class="btn" data-act="shared-mode" data-key="${spec.key}" data-mode="in_file">Use the sheet inside my file</button>` : `<button class="btn" data-act="shared-mode" data-key="${spec.key}" data-mode="shared">Use the SAP data MM60</button>`) : ""}
      <button class="link" data-act="shared-pick" data-key="${spec.key}">Use a different file for this run only</button></div>`;
  }
  const status = !info.loaded ? `<span class="pill warn">Not loaded</span>` : info.old ? `<span class="pill warn">${info.age_days} days old</span>` : `<span class="pill ok">✓ From SAP data</span>`;
  const line = spec.shared === "mm60_previous" ? (info.loaded ? `${esc(info.original_name)} · the file the current MM60 replaced` : "No older MM60 for this plant in the SAP data history yet")
    : info.loaded ? `${esc(info.original_name)} · ${esc(ageText(info))}` : `Load it on the SAP data page`;
  return `<div class="slot"><div class="grow"><strong>${esc(name)}</strong> ${status}<div>${line}</div></div>
    ${!info.loaded && spec.shared !== "mm60_previous" ? `<button class="btn" data-act="nav" data-to="sap">Open SAP data</button>` : ""}
    <button class="link" data-act="shared-pick" data-key="${spec.key}">Use a different file for this run only</button></div>`;
}
function fileBlock(spec) {
  const files = S.files[spec.key] || [];
  const list = files.map((f, i) => `<li><div class="grow"><div class="fname">${esc(f.name)}</div><div class="small muted">${esc(f.folder || f.path)}</div></div>
      <span class="pill ${f.recognised ? (f.fits ? "ok" : "warn") : "idle"}">${f.recognised ? (f.fits ? "✓ " : "") + esc(f.label) : "Not recognised"}</span>
      ${f.detail ? `<span class="small muted">${esc(f.detail)}</span>` : ""}
      <button class="btn" data-act="remove-file" data-key="${spec.key}" data-i="${i}" aria-label="Remove ${esc(f.name)}">Remove</button></li>`).join("");
  return `<div class="inblock"><div class="drop" data-drop="tool:${spec.key}"><p><strong>${esc(spec.label)}</strong>${spec.required ? "" : " (optional)"}<br>${S.mode.drop ? "Drop " + (spec.multiple ? "the files" : "the file") + " here, or browse." : "Use Browse to choose " + (spec.multiple ? "the files." : "the file.")}
      <span class="small">${(spec.accept || []).join(", ")}</span></p><button class="btn" data-act="browse" data-key="${spec.key}">Browse…</button></div>
    ${files.length ? `<ul class="files">${list}</ul>` : spec.empty_note ? `<p class="small muted" style="margin-top:8px">${esc(spec.empty_note)}</p>` : ""}
    ${files.some((f) => f.recognised && !f.fits) ? `<p class="note warn" style="margin-top:8px">A file here does not look like what this tool expects. You can still run: the script decides for itself.</p>` : ""}</div>`;
}
function optionField(o) {
  const v = S.opts[o.key];
  if (o.type === "checkbox") return `<label class="check"><input type="checkbox" data-opt="${o.key}"${v ? " checked" : ""}><span>${esc(o.label)}${o.help ? ` <span class="muted small">· ${esc(o.help)}</span>` : ""}</span></label>`;
  if (o.type === "choice") return `<div><div style="font-weight:500;margin-bottom:4px">${esc(o.label)}</div>${o.choices.map((c) => `<label class="check"><input type="radio" name="opt-${o.key}" data-opt="${o.key}" value="${esc(c.value)}"${v === c.value ? " checked" : ""}><span>${esc(c.label)}</span></label>`).join("")}</div>`;
  return `<label class="field"><span>${esc(o.label)}${o.required ? "" : ""}</span><input type="${o.type === "number" ? "number" : "text"}" ${o.type === "number" ? 'min="0" step="1" inputmode="numeric"' : ""} data-opt="${o.key}" value="${esc(v)}">${o.help ? `<small>${esc(o.help)}</small>` : ""}</label>`;
}
function optionsHtml() {
  const t = S.tool; let html = "";
  if (t.variant_option) {
    html += `<div style="margin-bottom:14px"><div style="font-weight:500;margin-bottom:6px">${esc(t.variant_label || "Choose")}</div><div class="seg" role="group" aria-label="${esc(t.variant_label || "Choose")}">${t.variants.map((v) =>
      `<button data-act="variant" data-v="${esc(v.value)}" aria-pressed="${S.opts[t.variant_option] === v.value}"${v.installed ? "" : " disabled"}>${esc(v.label)}${v.installed ? "" : " (not installed yet)"}</button>`).join("")}</div></div>`;
  }
  const plain = t.options.filter((o) => !o.group), groups = {};
  for (const o of t.options.filter((o) => o.group)) (groups[o.group] = groups[o.group] || []).push(o);
  if (plain.length) html += `<div class="stack">${plain.map(optionField).join("")}</div>`;
  for (const g in groups) html += `<h3 style="margin:18px 0 10px">${esc(g)}</h3><div class="opt-grid">${groups[g].map(optionField).join("")}</div>`;
  return html || `<p class="muted">Nothing to choose for this tool.</p>`;
}
function viewTool() {
  const t = S.tool;
  const head = `<div class="crumb"><button class="link" data-act="nav" data-to="home">Home</button> › <button class="link" data-act="nav-group" data-id="${t.group}">${esc(groupName(t.group))}</button> › ${esc(t.name)}</div>
    <div class="title-row"><h1>${esc(t.name)}</h1>${t.version ? `<span class="chip">${esc(t.version)}</span>` : ""}</div><p class="lead">${esc(t.what)}</p>`;
  if (!t.installed) return `<div class="wrap">${head}<div class="note warn"><strong>Not installed yet</strong> · add its folder and run the installer again.<br>The tool's script was not found in ${esc(t.tool_dir)}.</div></div>`;
  const side = `<aside><section class="card sidecard"><h3>Files for this tool</h3>
      ${t.open_files.map((f) => `<button class="btn" data-act="open" data-path="${esc(f.path)}"${f.exists ? "" : " disabled"}>${esc(f.label)}${f.exists ? "" : " (not there yet)"}</button>`).join("")}
      <button class="btn" data-act="open" data-path="${esc(t.tool_dir)}">Open tool folder</button></section>
    ${t.help ? `<section class="card sidecard"><h3>Good to know</h3><p class="small">${esc(t.help)}</p></section>` : ""}
    <section class="card sidecard"><h3>Recent runs</h3>${t.recent.length ? `<ul class="recent">${t.recent.map((r) => `<li>${esc(niceDate(r.time))} · ${esc(STATUS_WORD[r.status] || r.status)}<br><span class="muted">${esc((r.inputs || []).join(", ") || "no files")}</span></li>`).join("")}</ul>` : `<p class="small muted">Not run yet.</p>`}</section></aside>`;
  if (t.page === "top_bottom_sellers") {
    const wb = t.open_files[0];
    return `<div class="wrap">${head}<div class="tool-grid"><div>
      <section class="card"><h2 style="margin-bottom:12px">Each month: paste the Micros export</h2>
        <p class="paste-rule">Copy columns A to J only.<br>Paste at A7 on that month’s Input tab.</p>
        <p class="muted" style="margin:12px 0">The paste is done by hand. Nothing is overwritten: each month has its own Input tab.</p>
        <button class="btn primary big" data-act="open" data-path="${esc(wb.path)}"${wb.exists ? "" : " disabled"}>Open this year’s workbook</button>
        ${wb.exists ? "" : `<p class="small muted" style="margin-top:8px">This year’s workbook is not in the tool folder yet. Build it below.</p>`}</section>
      <section class="card"><div class="step-head"><h2>Build a new year workbook</h2></div>${optionsHtml()}
        <div style="margin-top:16px" id="checks">${checksHtml()}</div></section></div>${side}</div></div>`;
  }
  const inputs = t.inputs.map((spec) => spec.shared ? sharedSlot(spec) : fileBlock(spec)).join("");
  return `<div class="wrap">${head}<div class="tool-grid"><div>
    <section class="card"><div class="step-head"><span class="step-n">1</span><h2>Add your files</h2></div>${inputs || '<p class="muted">This tool needs no files.</p>'}</section>
    <section class="card"><div class="step-head"><span class="step-n">2</span><h2>Choose options</h2></div>${optionsHtml()}</section>
    <section class="card"><div class="step-head"><span class="step-n">3</span><h2>Check and run</h2></div><div id="checks">${checksHtml()}</div></section>
    </div>${side}</div></div>`;
}

// ---------- Running / Finished / Failed ----------
async function startRun() {
  const res = await api("start_run", S.tool.id, null, inputsPayload(), S.opts, false);
  if (res.error) { toast(res.error, 7000); runChecks(); return; }
  S.run = { id: res.run_id, tool: S.tool.id, name: S.tool.name, version: S.tool.version, progress: S.tool.progress || [], done: [], lines: [], cursor: 0,
            question: null, result: null, status: "Getting ready", pick: "", started: Date.now() };
  go({ name: "run" });
  pollRun();
}
async function pollRun() {
  const run = S.run;
  if (!run || run.result) return;
  let res;
  try { res = await api("poll", run.id, run.cursor); } catch (e) { setTimeout(pollRun, 600); return; }
  if (S.run !== run) return;
  run.cursor = res.cursor;
  let logChanged = false, qChanged = false;
  for (const ev of res.events) {
    if (ev.type === "line") { run.lines.push(ev); logChanged = true; }
    else if (ev.type === "prompt") { run.question = ev.card; run.pick = ""; qChanged = true; }
    else if (ev.type === "prompt_cleared") { run.question = null; qChanged = true; }
    else if (ev.type === "progress") { if (!run.done.includes(ev.index)) run.done.push(ev.index); qChanged = true; }
    else if (ev.type === "status") { run.status = ev.text; qChanged = true; }
  }
  if (res.question === null && run.question && !res.events.some((e) => e.type === "prompt")) { run.question = null; qChanged = true; }
  if (res.done && res.result) {
    run.result = res.result; run.question = null;
    await refresh();
    if (S.route.name === "run") render();
    else toast(`${run.name}: ${run.result.headline}`);
    return;
  }
  if (S.route.name === "run") {
    if (logChanged) appendLog();
    if (qChanged) { $("#runtop").innerHTML = runTopHtml(); $("#question").innerHTML = questionHtml(); focusQuestion(); markReply(); }
  }
  setTimeout(pollRun, res.events.length ? 120 : 250);
}
function lineHtml(ev) { return `<div class="l-${esc(ev.kind)}" data-n="${ev.n}">${esc(ev.text) || "&nbsp;"}</div>`; }
function appendLog() {
  const log = $("#log"); if (!log) return;
  const have = log.children.length, stick = log.scrollTop + log.clientHeight >= log.scrollHeight - 30;
  log.insertAdjacentHTML("beforeend", S.run.lines.slice(have).map(lineHtml).join(""));
  if (stick) log.scrollTop = log.scrollHeight;
}
function runTopHtml() {
  const run = S.run, waiting = !!run.question;
  const chips = run.progress.length ? run.progress.map((p, i) => `<span class="chip ${run.done.includes(i) ? "done" : ""}">${run.done.includes(i) ? "✓ " : ""}${esc(p)}</span>`).join("") : `<span class="chip">Running</span>`;
  return `<div class="row"><span class="pill ${waiting ? "warn" : "ok"}">${waiting ? "Waiting for your answer" : esc(run.status || "Running")}</span>
    <span class="grow"></span><button class="btn danger" data-act="stop">Stop run</button></div><div class="chips">${chips}</div>`;
}
function questionHtml() {
  const q = S.run.question;
  if (!q || q.kind === "reply") return "";
  const opts = (q.options || []).map((o) => `<button class="opt" data-act="pick" data-v="${esc(o.value)}" aria-pressed="${S.run.pick === o.value && S.run.picked}"><span class="key">${esc(o.value === "" ? "Enter" : o.value)}</span><span>${esc(o.label)}</span></button>`).join("");
  return `<section class="card question" aria-live="polite"><h2>${esc(q.title)}</h2>${q.help ? `<p class="muted" style="margin-top:4px">${esc(q.help)}</p>` : ""}
    ${opts ? `<div class="opts">${opts}</div>` : ""}
    <div class="row" style="margin-top:12px"><label class="grow"><span class="hide">Your answer</span><input type="text" id="qtext" placeholder="${opts ? "Or type an answer" : "Type your answer"}" value="${esc(S.run.pick)}" autocomplete="off"></label>
    <button class="btn primary" data-act="confirm">Confirm and continue</button></div>
    ${opts ? `<p class="small muted" style="margin-top:8px">Keyboard: press an option’s key to pick it, Enter to confirm.</p>` : ""}</section>`;
}
function focusQuestion() { const t = $("#qtext"); if (t) t.focus(); }
function markReply() { const r = $("#replybar"); if (r) r.classList.toggle("wanted", !!(S.run.question && S.run.question.kind === "reply")); if (S.run.question && S.run.question.kind === "reply") $("#reply").focus(); }
async function sendAnswer(text) {
  const run = S.run; if (!run || run.result) return;
  run.question = null; run.pick = ""; run.picked = false;
  $("#question").innerHTML = ""; $("#runtop").innerHTML = runTopHtml(); markReply();
  await api("answer", run.id, text);
}
function resultHtml() {
  const run = S.run, r = run.result, ok = r.status === "finished", bad = r.status === "failed";
  const outs = r.outputs.filter((o) => o.kind === "output" || o.kind === "log"), extra = r.outputs.filter((o) => o.kind === "data" || o.kind === "backup");
  const fileRow = (o) => `<li><div class="grow"><strong>${esc(o.name)}</strong> <span class="chip">${o.state === "new" ? "New" : "Updated"}</span><div class="small muted">${esc(o.folder)}</div></div>
      <button class="btn primary" data-act="open" data-path="${esc(o.path)}">${/\.xls/i.test(o.name) ? "Open in Excel" : "Open"}</button><button class="btn" data-act="reveal" data-path="${esc(o.path)}">Show in folder</button></li>`;
  return `<div class="banner ${ok ? "ok" : bad ? "bad" : "warn"}"><h2 class="grow">${ok ? "✓ " : ""}${esc(STATUS_WORD[r.status] || r.status)}${r.headline && r.headline !== "Finished" ? ": " + esc(r.headline.replace(/^Finished\.?\s*/, "")) : ""}</h2><span class="small">${esc(r.duration)} seconds</span></div>
    ${r.notices.map((n) => `<p class="note warn" style="margin-top:12px"><strong>From the script:</strong> ${esc(n)}</p>`).join("")}
    ${r.summary.length ? `<div class="totals">${r.summary.map((s) => `<div class="total"><div class="v">${esc(s.value)}</div><div class="k">${esc(s.label)}</div></div>`).join("")}</div>` : '<div style="height:16px"></div>'}
    <section class="card"><h2 style="margin-bottom:8px">Your files</h2>
      ${outs.length ? `<ul class="outs">${outs.map(fileRow).join("")}</ul>` : `<p class="muted">${ok ? "The script finished but no new or changed file was found." : "No file was written."}</p>`}
      ${extra.length ? `<details style="margin-top:10px"><summary class="muted">Also updated: ${extra.length} data or backup file${extra.length === 1 ? "" : "s"}</summary><ul class="outs">${extra.map(fileRow).join("")}</ul></details>` : ""}
      <div class="row" style="margin-top:14px">${r.log_path ? `<button class="btn" data-act="open" data-path="${esc(r.log_path)}">Open log</button>` : ""}
        ${bad ? `<button class="btn primary" data-act="copy-details">Copy details for Alain</button>` : ""}
        <button class="btn" data-act="open-tool" data-id="${esc(r.tool)}">Back to ${esc(r.tool_name)}</button></div></section>
    ${ok && r.help ? `<section class="card"><h2 style="margin-bottom:6px">Next</h2><p>${esc(r.help)}</p></section>` : ""}
    ${r.review.length ? `<section class="card"><h2 style="margin-bottom:6px">Worth a look</h2><ul class="looks">${r.review.map((x) => `<li><button data-act="jump" data-n="${x.line}">${esc(x.text)}</button></li>`).join("")}</ul></section>` : ""}`;
}
function viewRun() {
  const run = S.run;
  if (!run) return `<div class="wrap"><p class="muted">Nothing is running.</p></div>`;
  const head = `<div class="crumb"><button class="link" data-act="nav" data-to="home">Home</button> › <button class="link" data-act="open-tool" data-id="${run.tool}">${esc(run.name)}</button> › ${run.result ? "Result" : "Running"}</div>
    <div class="title-row"><h1>${esc(run.name)}</h1>${run.version ? `<span class="chip">${esc(run.version)}</span>` : ""}</div>`;
  const log = `<section class="logbox" style="margin-top:16px"><div class="logbar"><span class="grow">${run.result ? "Full log" : "Live log"}</span><button class="btn" data-act="copy-log">Copy</button></div>
    <div class="log" id="log" tabindex="0" aria-label="Log">${run.lines.map(lineHtml).join("")}</div>
    ${run.result ? "" : `<form class="reply" id="replybar" data-form="reply"><label class="grow"><span class="hide">Reply to the script</span><input type="text" id="reply" placeholder="Reply to the script, then press Enter" autocomplete="off"></label><button class="btn" type="submit">Send</button></form>`}</section>`;
  if (run.result) return `<div class="wrap">${head}<div style="margin-top:16px">${resultHtml()}</div>${log}${historyTable(run.tool)}</div>`;
  return `<div class="wrap">${head}<div id="runtop" style="margin-top:16px">${runTopHtml()}</div><div id="question">${questionHtml()}</div>${log}</div>`;
}
function afterRunRender() { const log = $("#log"); if (log) log.scrollTop = log.scrollHeight; if (S.run && !S.run.result) { focusQuestion(); markReply(); } }
function historyTable(tool) {
  return `<section class="card" style="margin-top:16px"><div class="row"><h2 class="grow">Run history</h2><button class="link" data-act="nav" data-to="history">See all</button></div><div id="mini-hist" data-tool="${esc(tool)}"><p class="muted small">Loading…</p></div></section>`;
}
function histRows(rows, withTool) {
  if (!rows.length) return `<p class="muted">No runs yet.</p>`;
  return `<div class="tablewrap"><table><thead><tr><th>When</th>${withTool ? "<th>Tool</th>" : ""}<th>Result</th><th>Who</th><th>Files in</th><th>Files out</th><th>Took</th><th>Log</th></tr></thead><tbody>${rows.map((r) => `<tr>
    <td class="num">${esc(niceDate(r.time))}</td>${withTool ? `<td>${esc(r.tool_name || r.tool)} <span class="muted small">${esc(r.version || "")}</span></td>` : ""}
    <td>${statusPill(r.status)}</td><td>${esc(r.user)}</td><td class="small">${esc((r.inputs || []).join(", "))}</td>
    <td class="small">${(r.outputs || []).map((p) => `<button class="link" data-act="open" data-path="${esc(p)}">${esc(p.split(/[\\/]/).pop())}</button>`).join("<br>")}</td>
    <td class="num">${esc(r.duration)}s</td><td>${r.log ? `<button class="link" data-act="open" data-path="${esc(r.log)}">Open</button>` : ""}</td></tr>`).join("")}</tbody></table></div>`;
}
async function fillMiniHist() { const el = $("#mini-hist"); if (!el) return; const rows = await api("history_rows", el.dataset.tool, null); el.innerHTML = histRows(rows.slice(0, 8), false); }

// ---------- SAP data ----------
function viewSap() {
  const sap = S.sap || S.st.sap;
  const card = (k, title, blurb) => {
    const i = sap[k], hist = i.history || [];
    const facts = i.loaded ? `<dl class="kv"><dt>File</dt><dd>${esc(i.original_name)}</dd>
        <dt>Exported</dt><dd>${esc(niceDate(i.exported))} <span class="muted small">(${esc(i.exported_source)})</span></dd>
        <dt>Age</dt><dd>${i.old ? `<span class="pill warn">${i.age_days} days old</span> older than your ${i.warn_days}-day warning` : `<span class="pill ok">✓ ${i.age_days === 0 ? "Today" : i.age_days + " days old"}</span>`}</dd>
        <dt>Loaded</dt><dd>${esc(niceDate(i.loaded_on))} by ${esc(i.loaded_by || "")}</dd><dt>Rows</dt><dd class="num">${esc((i.rows ?? "").toLocaleString())}</dd>
        ${i.plants && i.plants.length ? `<dt>Plant</dt><dd>${esc(i.plants.join(", "))}</dd>` : ""}
        ${i.locations && i.locations.length ? `<dt>Storage locations</dt><dd>${esc(i.locations.join(", "))}<div class="small muted">Check none are missing from your merged file.</div></dd>` : ""}</dl>
        <div class="row"><button class="btn" data-act="open" data-path="${esc(i.path)}">Open in Excel</button></div>` : `<p class="note warn" style="margin:12px 0"><strong>Not loaded yet.</strong> Add your SAP export below.</p>`;
    return `<section class="card"><div class="row"><h2 class="grow">${title}</h2></div><p class="muted" style="margin-top:4px">${blurb}</p>${facts}
      <div class="drop" data-drop="sap:${k}" style="margin-top:14px"><p><strong>${i.loaded ? "Replace with a newer export" : "Load the export"}</strong><br>${S.mode.drop ? "Drop the SAP export (.xlsx) here, or browse." : "Use Browse to choose the SAP export (.xlsx)."} The column headers are checked first.</p><button class="btn primary" data-act="sap-browse" data-kind="${k}">Browse…</button></div>
      ${hist.length ? `<details style="margin-top:12px"><summary>Earlier files (${hist.length}) · nothing is deleted</summary><ul class="outs">${hist.map((h) => `<li><div class="grow">${esc(h.name)}</div><button class="btn" data-act="reveal" data-path="${esc(h.path)}">Show in folder</button></li>`).join("")}</ul></details>` : ""}</section>`;
  };
  return `<div class="wrap"><h1>SAP data</h1><p class="lead">One MM60 per plant and one kitchen MI20, loaded once here and replaced when you have a newer export. Every tool that needs them reads them from here.</p>
    <p class="note info" style="margin-bottom:16px">An MM60 is recognised by its Plant column: <strong>${esc(PL().hotel)}</strong> is the ${esc(PL().names[PL().hotel] || "hotel")}, <strong>${esc(PL().cc)}</strong> is the ${esc(PL().names[PL().cc] || "conference centre")}. Drop it on either MM60 card and it goes to the right one.</p>
    ${card("mm60", esc(SAP_SHORT.mm60), "Used by OE Monthly Run, the MM60 Cost Comparator, and Theoretical CoS for hotel bar sales.")}
    ${card("mm60_cc", esc(SAP_SHORT.mm60_cc), "Used by Theoretical CoS for conference centre sales and by the MM60 Cost Comparator.")}
    ${card("mi20_kitchen", "Kitchen MI20 (your merged monthly file)", "Used by the Coffee Shop, Breakfast and Canteen trackers. Taken exactly as it is: no merging, re-sorting or reformatting.")}
    <p class="muted small" style="margin-top:14px">Other SAP exports (MB52, MB51, variance reports, Micros) are added on each tool’s page as before.</p></div>`;
}
async function openSap() { S.sap = await api("sap_status"); go({ name: "sap" }); }
async function sapReplace(kind, path) {
  const chk = await api("sap_inspect", kind, path);
  const name = path.split(/[\\/]/).pop();
  if (!chk.ok) return modal(`<h2>That file was not loaded</h2><p class="note bad">${esc(chk.reason)}</p><p style="margin-top:10px">The current file is still in place.</p><div class="row end"><button class="btn primary" data-act="close-modal">OK</button></div>`);
  kind = chk.kind || kind;
  const cur = (S.sap || S.st.sap)[kind];
  S.pendingSap = { kind, path };
  modal(`<h2>${cur.loaded ? "Replace" : "Load"} ${esc(SAP_SHORT[kind])}?</h2>
    <dl class="kv"><dt>New file</dt><dd>${esc(name)}</dd><dt>Rows</dt><dd class="num">${chk.rows.toLocaleString()}</dd>
    ${chk.plants.length ? `<dt>Plant</dt><dd>${esc(chk.plants.join(", "))}</dd>` : ""}${chk.locations.length ? `<dt>Storage locations</dt><dd>${esc(chk.locations.join(", "))}</dd>` : ""}</dl>
    <p class="muted">${cur.loaded ? `The current file (${esc(cur.original_name)}) moves to the history folder. Nothing is deleted.` : "The column headers look right."}</p>
    <div class="row end"><button class="btn" data-act="close-modal">Cancel</button><button class="btn primary" data-act="sap-confirm">${cur.loaded ? "Replace" : "Load"}</button></div>`);
}

// ---------- History ----------
function viewHistory() {
  const tools = S.st.tools.map((t) => `<option value="${t.id}"${S.hist.tool === t.id ? " selected" : ""}>${esc(t.name)}</option>`).join("");
  return `<div class="wrap"><h1>Run history</h1><p class="lead">Every run, newest first.</p>
    <div class="card"><div class="row" style="margin-bottom:14px"><label class="field" style="width:280px"><span>Tool</span><select id="h-tool"><option value="">All tools</option>${tools}</select></label>
      <label class="field" style="width:200px"><span>Month</span><input type="month" id="h-month" value="${esc(S.hist.month)}"></label>
      <button class="btn" data-act="hist-clear" style="align-self:flex-end">Show everything</button></div><div id="hist"><p class="muted">Loading…</p></div></div></div>`;
}
async function fillHistory() { const el = $("#hist"); if (!el) return; el.innerHTML = histRows(await api("history_rows", S.hist.tool, S.hist.month), true); }

// ---------- Settings ----------
function viewSettings() {
  const s = S.st.settings, folders = s.data_folders || {};
  const seen = new Set(), rows = [];
  for (const t of S.st.tools) {
    const f = S.toolFolders && S.toolFolders[t.id]; if (!f || seen.has(f.folder)) continue; seen.add(f.folder);
    rows.push(`<tr><td>${esc(f.names.join(", "))}</td><td class="small" style="overflow-wrap:anywhere">${esc(f.dir)}${folders[f.folder] ? ' <span class="pill warn">Custom</span>' : ""}</td>
      <td style="white-space:nowrap"><button class="btn" data-act="folder-pick" data-folder="${esc(f.folder)}">Change…</button> ${folders[f.folder] ? `<button class="btn" data-act="folder-reset" data-folder="${esc(f.folder)}">Use default</button>` : ""} <button class="btn" data-act="open" data-path="${esc(f.dir)}">Open</button></td></tr>`);
  }
  const checklist = S.st.tools.map((t) => `<label class="check"><input type="checkbox" data-checklist="${t.id}"${(s.month_end_checklist || []).includes(t.id) ? " checked" : ""}><span>${esc(t.name)}</span></label>`).join("");
  const setup = S.setup ? `<div class="tablewrap" style="margin-top:12px"><table><tbody>${S.setup.map((r) => `<tr><td>${esc(r.what)}</td><td>${r.ok ? '<span class="pill ok">✓ OK</span>' : '<span class="pill bad">Problem</span>'}</td><td class="small" style="overflow-wrap:anywhere">${esc(r.value)}</td></tr>`).join("")}</tbody></table></div>` : "";
  return `<div class="wrap"><h1>Settings &amp; data folders</h1><p class="lead">Your data lives in ${esc(S.st.app.data_root)} and is never replaced by an update.</p>
    <section class="card"><h2 style="margin-bottom:12px">SAP data</h2>
      <label class="field" style="max-width:340px"><span>Warn me when MM60 or MI20 is older than (days)</span><input type="number" id="set-age" min="1" max="365" value="${esc(s.sap_age_warning_days)}"><small>A warning only. It never blocks a run.</small></label>
      <label class="check" style="margin-top:10px"><input type="checkbox" id="set-pr"${s.pr_tool_copy_mm60 ? " checked" : ""}><span>Also copy MM60 to the PR Tool data folder when I replace it</span></label>
      <div class="row"><label class="field grow"><span>PR Tool data folder</span><input type="text" id="set-prdir" value="${esc(s.pr_tool_data_folder || "")}" placeholder="Folder that holds the PR Tool's mm60*.xlsx"></label><button class="btn" data-act="pr-pick" style="align-self:flex-end">Browse…</button></div></section>
    <section class="card"><h2 style="margin-bottom:6px">Data folders</h2><p class="muted" style="margin-bottom:10px">Point a tool at a shared folder (for example a network share) so everyone uses the same item map. Copy the tool's folder there first.</p>
      <div class="tablewrap"><table><thead><tr><th>Tool</th><th>Folder</th><th></th></tr></thead><tbody>${rows.join("") || '<tr><td colspan="3" class="muted">Loading…</td></tr>'}</tbody></table></div></section>
    <section class="card"><h2 style="margin-bottom:6px">Month-end checklist</h2><p class="muted" style="margin-bottom:6px">Tools shown on Home and ticked when they finish in the current month.</p><div class="opt-grid">${checklist}</div></section>
    <section class="card"><h2 style="margin-bottom:6px">Help</h2><div class="row"><button class="btn" data-act="open-logs">Open logs folder</button><button class="btn primary" data-act="check-setup">Check my setup</button></div>${setup}</section></div>`;
}
async function openSettings() {
  S.toolFolders = {};
  for (const t of S.st.tools) {
    const v = await api("tool_view", t.id);
    const folder = v.tool_dir.split(/[\\/]/).pop();
    const key = Object.keys(S.toolFolders).find((k) => S.toolFolders[k].dir === v.tool_dir);
    if (key) { S.toolFolders[key].names.push(t.name); S.toolFolders[t.id] = S.toolFolders[key]; }
    else S.toolFolders[t.id] = { folder: (S.st.settings.data_folders && Object.keys(S.st.settings.data_folders).find((k) => S.st.settings.data_folders[k] === v.tool_dir)) || folder, dir: v.tool_dir, names: [t.name] };
  }
  go({ name: "settings" });
}
async function saveSettings(changes) { await api("save_settings", changes); await refresh(); if (S.route.name === "settings") await openSettings(); }

// ---------- dropped files ----------
let dropZone = null;
window.ccsDropped = async function (paths) {
  $$(".dropveil").forEach((e) => e.remove());
  const zone = dropZone; dropZone = null;
  if (zone && zone.startsWith("tool:") && S.route.name === "tool") return addFiles(zone.slice(5), paths);
  if (zone && zone.startsWith("sap:") && S.route.name === "sap") return sapReplace(zone.slice(4), paths[0]);
  if (S.route.name === "tool" && firstFileKey()) return addFiles(firstFileKey(), paths);
  const res = await api("suggest", paths);
  const files = res.files.map((f) => `<li><div class="grow"><strong>${esc(f.name)}</strong></div><span class="pill ${f.kind === "unknown" ? "idle" : "ok"}">${esc(f.label)}</span></li>`).join("");
  S.dropped = paths;
  const tools = res.tools.map((t) => `<button class="btn primary" data-act="suggest-open" data-id="${t.id}">${t.id === "sap_data" ? "Load on the SAP data page" : "Open " + esc(t.name)}</button>`).join("");
  modal(`<h2>What would you like to do with ${paths.length === 1 ? "this file" : "these files"}?</h2><ul class="files">${files}</ul>
    <div class="row" style="margin-top:14px">${tools || '<p class="muted">The app does not recognise these files. Open the tool you want and add them there.</p>'}</div>
    <p class="small muted" style="margin-top:10px">Nothing runs until you press Run on the tool’s page.</p><div class="row end"><button class="btn" data-act="close-modal">Cancel</button></div>`);
};
document.addEventListener("dragover", (e) => { e.preventDefault(); const z = e.target.closest && e.target.closest("[data-drop]"); $$(".drop.over").forEach((d) => d.classList.remove("over")); if (z) z.classList.add("over"); });
document.addEventListener("dragleave", (e) => { if (!e.relatedTarget) $$(".drop.over").forEach((d) => d.classList.remove("over")); });
document.addEventListener("drop", (e) => {
  e.preventDefault();
  const z = e.target.closest && e.target.closest("[data-drop]");
  dropZone = z ? z.dataset.drop : null;
  $$(".drop.over").forEach((d) => d.classList.remove("over"));
  if (!S.mode.drop) toast("In this window, use Browse to choose files.");
});

// ---------- events ----------
document.addEventListener("click", async (e) => {
  const b = e.target.closest("[data-act]"); if (!b || b.disabled) return;
  const act = b.dataset.act, d = b.dataset;
  if (act === "nav") { if (d.to === "sap") return openSap(); if (d.to === "settings") return openSettings(); go({ name: d.to }); if (d.to === "history") fillHistory(); }
  else if (act === "nav-group") go({ name: "group", group: d.id });
  else if (act === "clear-search") { S.search = ""; render(); }
  else if (act === "open-tool") { closeModal(); openTool(d.id); }
  else if (act === "open") { const r = await api("open_path", d.path); if (!r.ok) toast(r.reason || "Could not open that."); }
  else if (act === "reveal") { const r = await api("reveal", d.path); if (!r.ok) toast(r.reason || "Could not show that."); }
  else if (act === "browse") { const spec = S.tool.inputs.find((i) => i.key === d.key); addFiles(d.key, await api("pick_files", !!spec.multiple, spec.accept || [".xlsx"])); }
  else if (act === "remove-file") { S.files[d.key].splice(+d.i, 1); render(); runChecks(); }
  else if (act === "variant") { S.opts[S.tool.variant_option] = d.v; render(); runChecks(); }
  else if (act === "shared-pick") { const p = await api("pick_files", false, [".xlsx"]); if (p.length) { S.shared[d.key] = { path: p[0], name: p[0].split(/[\\/]/).pop() }; render(); runChecks(); } }
  else if (act === "shared-reset") { delete S.shared[d.key]; render(); runChecks(); }
  else if (act === "shared-mode") { if (d.mode === "shared") delete S.shared[d.key]; else S.shared[d.key] = { mode: d.mode }; render(); runChecks(); }
  else if (act === "recheck") runChecks();
  else if (act === "run") startRun();
  else if (act === "stop") modal(`<h2>Stop this run?</h2><p>The script is ended straight away. Check the output folder afterwards: a file may be half written.</p><div class="row end"><button class="btn" data-act="close-modal">Keep running</button><button class="btn primary" data-act="stop-yes">Stop run</button></div>`);
  else if (act === "stop-yes") { closeModal(); await api("stop", S.run.id); }
  else if (act === "pick") { S.run.pick = d.v; S.run.picked = true; if (d.v === "") return sendAnswer(""); $("#question").innerHTML = questionHtml(); focusQuestion(); }
  else if (act === "confirm") sendAnswer($("#qtext").value);
  else if (act === "copy-log") copyText(S.run.lines.map((l) => l.text).join("\n"));
  else if (act === "copy-details") copyText(S.run.result.details);
  else if (act === "jump") { const el = $(`#log [data-n="${d.n}"]`); if (el) { $$("#log .hit").forEach((h) => h.classList.remove("hit")); el.classList.add("hit"); el.scrollIntoView({ block: "center" }); $("#log").focus(); } }
  else if (act === "close-modal") closeModal();
  else if (act === "sap-browse") { const p = await api("pick_files", false, [".xlsx"]); if (p.length) sapReplace(d.kind, p[0]); }
  else if (act === "sap-confirm") {
    const { kind, path } = S.pendingSap; closeModal();
    const r = await api("sap_replace", kind, path);
    if (!r.ok) return modal(`<h2>That file was not loaded</h2><p class="note bad">${esc(r.reason)}</p><div class="row end"><button class="btn primary" data-act="close-modal">OK</button></div>`);
    toast(SAP_SHORT[r.kind || kind] + " loaded." + (r.pr_tool ? " " + r.pr_tool : ""));
    await refresh(); openSap();
  }
  else if (act === "suggest-open") { closeModal(); if (d.id === "sap_data") { await openSap(); toast("Use Browse on the file you want to replace."); } else openTool(d.id, S.dropped); }
  else if (act === "hist-clear") { S.hist = { tool: "", month: "" }; render(); fillHistory(); }
  else if (act === "open-logs") api("open_logs");
  else if (act === "check-setup") { S.setup = await api("check_setup"); render(); }
  else if (act === "pr-pick") { const p = await api("pick_folder"); if (p.length) saveSettings({ pr_tool_data_folder: p[0] }); }
  else if (act === "folder-pick") { const p = await api("pick_folder"); if (p.length) { const f = Object.assign({}, S.st.settings.data_folders); f[d.folder] = p[0]; saveSettings({ data_folders: f }); } }
  else if (act === "folder-reset") { const f = Object.assign({}, S.st.settings.data_folders); delete f[d.folder]; saveSettings({ data_folders: f }); }
});
document.addEventListener("input", (e) => {
  const t = e.target;
  if (t.id === "search") { S.search = t.value; const pos = t.selectionStart; render(); const n = $("#search"); n.focus(); n.setSelectionRange(pos, pos); }
  else if (t.dataset.opt) { S.opts[t.dataset.opt] = t.type === "checkbox" ? t.checked : t.value; if (t.type === "radio") render(); runChecks(); }
  else if (t.id === "qtext") { S.run.pick = t.value; S.run.picked = false; }
});
document.addEventListener("change", (e) => {
  const t = e.target;
  if (t.id === "h-tool") { S.hist.tool = t.value; fillHistory(); }
  else if (t.id === "h-month") { S.hist.month = t.value; fillHistory(); }
  else if (t.id === "set-age") saveSettings({ sap_age_warning_days: t.value });
  else if (t.id === "set-pr") saveSettings({ pr_tool_copy_mm60: t.checked });
  else if (t.id === "set-prdir") saveSettings({ pr_tool_data_folder: t.value });
  else if (t.dataset.checklist) { const ids = $$("[data-checklist]").filter((c) => c.checked).map((c) => c.dataset.checklist); saveSettings({ month_end_checklist: ids }); }
});
document.addEventListener("submit", (e) => { if (e.target.dataset.form === "reply") { e.preventDefault(); const i = $("#reply"); const v = i.value; i.value = ""; sendAnswer(v); } });
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && !$("#modal").hidden) return closeModal();
  const q = S.run && !S.run.result && S.route.name === "run" ? S.run.question : null;
  if (!q || q.kind === "reply" || !$("#modal").hidden) return;
  const inText = e.target.id === "qtext", empty = !inText || e.target.value === "";
  if (e.key === "Enter" && (inText || e.target === document.body || e.target.id === "main")) { e.preventDefault(); return sendAnswer($("#qtext").value); }
  if (e.target.id === "reply" || e.ctrlKey || e.altKey || e.metaKey || e.key.length !== 1) return;
  // a single key picks an option only when it is the whole answer so far (so "R1" or a SAP code can still be typed)
  const hit = (q.options || []).find((o) => o.value.toUpperCase() === e.key.toUpperCase());
  const longer = (q.options || []).some((o) => o.value.length > 1 && o.value.toUpperCase().startsWith(e.key.toUpperCase()));
  if (hit && empty && !longer && !inText) { e.preventDefault(); S.run.pick = hit.value; S.run.picked = true; $("#question").innerHTML = questionHtml(); focusQuestion(); }
});
const observer = new MutationObserver(() => { if ($("#mini-hist") && !$("#mini-hist").dataset.done) { $("#mini-hist").dataset.done = "1"; fillMiniHist(); } });

// ---------- start ----------
(async function start() {
  await pyReady();
  try { S.mode = await api("mode"); } catch (e) { /* keep defaults */ }
  await refresh();
  observer.observe($("#main"), { childList: true, subtree: true });
  render();
  if (TOKEN) setInterval(() => api("mode").catch(() => {}), 20000);   // tells the local server the page is still open
})();
