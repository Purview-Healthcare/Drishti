import React, { useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import * as XLSX from "xlsx";

/* =========================================================================
   SLA & KPI Monitor — scope-based multi-client, multi-user console
   Each client is set up with only the scopes of work it deals with
   (AR / EV / Submission / Payment Posting / Credentialing / PCC / SLA);
   supervisors see only those. Formulas per the source file.
   ========================================================================= */

const isNum = (v) => typeof v === "number" && !isNaN(v);
const uid = () => Math.random().toString(36).slice(2, 9);
const nowIso = () => new Date().toISOString();
const n = (v) => { if (v === "" || v == null) return null; const x = Number(String(v).replace(/[$,%\s]/g, "")); return isNaN(x) ? null : x; };
const fmtDateTime = (iso) => { try { return new Date(iso).toLocaleString(undefined, { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" }); } catch { return iso; } };
async function sha256(str) {
  try { const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(str)); return [...new Uint8Array(b)].map((x) => x.toString(16).padStart(2, "0")).join(""); }
  catch { let h = 5381; for (let i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) >>> 0; return "f" + h.toString(16); }
}

/* --------------------------- persistence ------------------------------- */
const K = { users: "slamon:users:v3", clients: "slamon:clients:v3", audit: "slamon:audit:v3", session: "slamon:session:v3", questions: "slamon:questions:v3", qseed: "slamon:qseed:v3", feedback: "slamon:feedback:v3" };
const clientKey = (id) => `slamon:client:${id}:v3`;
/* Storage: inside Claude the app uses Claude's shared storage. Anywhere else (GitHub Pages, a laptop)
   it falls back to the browser's own database (IndexedDB), which holds far more than localStorage.
   Data saved this way stays on that one computer and browser; it is never sent anywhere. */
const LOCAL_DB = "drishti-store", LOCAL_TABLE = "kv";
let localDbPromise = null;
function openLocalDb() {
  if (localDbPromise) return localDbPromise;
  localDbPromise = new Promise((resolve, reject) => {
    try {
      const req = indexedDB.open(LOCAL_DB, 1);
      req.onupgradeneeded = () => { if (!req.result.objectStoreNames.contains(LOCAL_TABLE)) req.result.createObjectStore(LOCAL_TABLE); };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    } catch (e) { reject(e); }
  });
  return localDbPromise;
}
const localStore = {
  async get(key) { const db = await openLocalDb(); return new Promise((res, rej) => { const r = db.transaction(LOCAL_TABLE, "readonly").objectStore(LOCAL_TABLE).get(key); r.onsuccess = () => res(r.result == null ? null : { key, value: r.result }); r.onerror = () => rej(r.error); }); },
  async set(key, value) { const db = await openLocalDb(); return new Promise((res, rej) => { const t = db.transaction(LOCAL_TABLE, "readwrite"); t.objectStore(LOCAL_TABLE).put(value, key); t.oncomplete = () => res({ key, value }); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error); }); },
  async delete(key) { const db = await openLocalDb(); return new Promise((res, rej) => { const t = db.transaction(LOCAL_TABLE, "readwrite"); t.objectStore(LOCAL_TABLE).delete(key); t.oncomplete = () => res({ key, deleted: true }); t.onerror = () => rej(t.error); }); },
};
/* ---------------- shared server: a Google Sheet behind an Apps Script web app ----------------
   When config.js sets window.DRISHTI_API, every account, client and number lives in that one sheet,
   so the whole team sees the same data from any computer, and who sees what is decided on the server.
   Without it the app falls back to this browser's own storage (one computer only). */
const API_URL = (typeof window !== "undefined" && window.DRISHTI_API && String(window.DRISHTI_API).trim()) || "";
const REMOTE = /^https:\/\/|^http:\/\/(localhost|127\.0\.0\.1)[:/]/.test(API_URL);
const TOKEN_KEY = "drishti:token:v1";
let memToken = "";
const tokGet = () => { try { return localStorage.getItem(TOKEN_KEY) || memToken; } catch { return memToken; } };
const tokSet = (t) => { memToken = t || ""; try { if (t) localStorage.setItem(TOKEN_KEY, t); else localStorage.removeItem(TOKEN_KEY); } catch {} };
const emit = (name, detail) => { try { window.dispatchEvent(new CustomEvent(name, { detail })); } catch {} };
const apiErr = (message, code) => { const e = new Error(message); e.code = code; return e; };
async function api(a, body = {}) {
  let res;
  // text/plain keeps the browser from sending a pre-check request, which Apps Script does not answer
  try { res = await fetch(API_URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify({ ...body, a, t: tokGet() }), redirect: "follow" }); }
  catch { throw apiErr("Drishti cannot reach its server. Check your internet connection and try again.", "network"); }
  let j; try { j = await res.json(); } catch { throw apiErr("The Drishti server sent back something unexpected. Try again in a minute.", "server"); }
  if (!j || !j.ok) { const e = apiErr((j && j.error) || "Something went wrong.", j && j.code); if (e.code === "auth") emit("drishti:auth", e.message); throw e; }
  return j;
}
// report rows never leave the browser: only the figures worked out from them are saved
function stripPeriod(p) {
  const c = JSON.parse(JSON.stringify(p || {}));
  for (const r of Object.values(c.reports || {})) if (r && typeof r === "object") { delete r.raw; delete r.rows_; delete r.columns; }
  delete c._stub; delete c._resave;
  return c;
}
// a client's own settings (scopes, targets, remembered columns); its weeks and months are saved one by one
const settingsOf = (rec) => { const { periods, ownerId, ownerName, archived, archivedAt, archivedBy, ...rest } = rec || {}; return rest; };
const hashStr = (s) => { let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return (h >>> 0).toString(36) + "." + s.length; };
const isClientKey = (k) => /^slamon:client:.+:v3$/.test(k);
const idOfClientKey = (k) => k.slice("slamon:client:".length, -3);
// what the server already holds for each client, so a save sends only the settings or periods that changed
const known = new Map();
function rememberClient(rec) {
  if (!rec || !rec.id) return;
  const periods = new Map();
  for (const p of rec.periods || []) {
    if (p._stub) { periods.set(p.id, { stub: true }); continue; }
    if (p._resave) { periods.set(p.id, { hash: "", rhash: "" }); continue; }
    const sp = stripPeriod(p); periods.set(p.id, { hash: hashStr(JSON.stringify(sp)), rhash: hashStr(JSON.stringify(sp.reports || {})) });
  }
  known.set(rec.id, { settings: hashStr(JSON.stringify(settingsOf(rec))), periods });
}
function rememberPeriod(clientId, p) {
  const k = known.get(clientId); if (!k || !p) return;
  const sp = stripPeriod(p); k.periods.set(p.id, { hash: hashStr(JSON.stringify(sp)), rhash: hashStr(JSON.stringify(sp.reports || {})) });
}
async function apiRetry(a, body) {
  try { return await api(a, body); }
  catch (e) { if (e.code !== "network" && e.code !== "server") throw e; await new Promise((r) => setTimeout(r, 1500)); return api(a, body); }
}
// Send the settings if they changed, and every week or month that changed, each with the rows for the
// Google Sheet's KPI History tab. Report Details rows go only when that period's reports changed.
async function syncClient(rec) {
  const k = known.get(rec.id) || { settings: null, periods: new Map() }; known.set(rec.id, k);
  const sh = hashStr(JSON.stringify(settingsOf(rec)));
  if (sh !== k.settings) { await apiRetry("client.put", { value: settingsOf(rec) }); k.settings = sh; }
  const ids = new Set();
  for (const p of rec.periods || []) {
    ids.add(p.id); if (p._stub) continue;
    const sp = stripPeriod(p), h = hashStr(JSON.stringify(sp)), was = k.periods.get(p.id);
    if (was && !was.stub && was.hash === h) continue;
    const rh = hashStr(JSON.stringify(sp.reports || {}));
    await apiRetry("period.put", { clientId: rec.id, period: sp, type: periodType(sp), summary: periodSummary(rec, sp), lite: periodLite(sp), history: historyRows(rec, sp), breakdowns: !was || was.stub || was.rhash !== rh ? breakdownRows(sp) : null });
    k.periods.set(p.id, { hash: h, rhash: rh });
  }
  for (const pid of [...k.periods.keys()]) if (!ids.has(pid)) { await apiRetry("period.delete", { clientId: rec.id, id: pid }); k.periods.delete(pid); }
}
const remoteLatest = new Map(); // unsaved client edits, newest first in line
const writeLines = new Map();   // one save at a time per client, so an older save can never land after a newer one
function queueClientPut(key, json) {
  remoteLatest.set(key, json);
  let line = writeLines.get(key);
  if (!line) { line = { next: null, running: null }; writeLines.set(key, line); }
  line.next = json;
  if (!line.running) line.running = (async () => {
    try {
      while (line.next != null) { const body = line.next; line.next = null; await syncClient(JSON.parse(body)); }
      emit("drishti:saved");
    } catch (e) {
      if (e.code !== "auth") emit("drishti:problem", "Your latest change was not saved. " + (e.message || "") + " Make the change again once you are back online.");
      throw e;
    } finally {
      line.running = null; const left = line.next; writeLines.delete(key);
      if (left != null) queueClientPut(key, left).catch(() => {}); else remoteLatest.delete(key);
    }
  })();
  return line.running;
}
const pendingSaves = () => writeLines.size > 0;
const remoteStore = {
  async get(key) {
    if (writeLines.has(key)) return { key, value: remoteLatest.get(key) }; // our own unsaved edit is newer than the server copy
    let v;
    if (key === K.users) v = (await api("users.list")).users;
    else if (key === K.clients) v = (await api("clients.list")).clients;
    else if (isClientKey(key)) {
      v = (await api("client.get", { id: idOfClientKey(key) })).value;
      if (v) {
        v.periods = sortPeriods(v.periods); rememberClient(v);
        // periods moved over from the first version are saved once more so their history rows get written
        if (v.periods.some((p) => p._resave)) setTimeout(() => queueClientPut(key, JSON.stringify(v)).catch(() => {}), 0);
      }
    }
    else if (key === K.audit) v = (await api("audit.list")).value;
    else if (key === K.feedback) v = (await api("feedback.list")).value;
    else if (key === K.questions || key === K.qseed) { const s = (await api("kv.get", { key })).value; return s == null || s === "" ? null : { key, value: s }; }
    else return null;
    return v == null ? null : { key, value: JSON.stringify(v) };
  },
  async set(key, value) {
    if (isClientKey(key)) return queueClientPut(key, value);
    if (key === K.questions || key === K.qseed) return api("kv.set", { key, value });
    // accounts, the client list, the audit log and feedback each have their own server action
    return null;
  },
  async delete(key) { if (isClientKey(key)) { writeLines.delete(key); remoteLatest.delete(key); return api("client.delete", { id: idOfClientKey(key) }); } return null; },
};
const store = () => (REMOTE ? remoteStore : typeof window !== "undefined" && window.storage && typeof window.storage.get === "function" ? window.storage : localStore);
async function sGet(key, shared = true) { try { const r = await store().get(key, shared); return r ? JSON.parse(r.value) : null; } catch (e) { if (REMOTE && e && e.code !== "auth" && e.code !== "forbidden") emit("drishti:problem", e.message); return null; } }
async function sSet(key, value, shared = true) { try { await store().set(key, JSON.stringify(value), shared); return true; } catch { return false; } }
async function appendAudit(entry) {
  if (REMOTE) { try { await api("audit.add", { entry }); } catch {} return; }
  const log = (await sGet(K.audit)) || []; log.unshift({ id: uid(), ts: nowIso(), ...entry }); await sSet(K.audit, log.slice(0, 600));
}
// An older week or month arrives as a short copy; this fetches the whole period when someone picks it.
async function loadFullPeriod(clientId, periodId) {
  if (!REMOTE) return null;
  try { const p = (await api("period.get", { clientId, id: periodId })).value; if (p) rememberPeriod(clientId, p); return p || null; }
  catch (e) { if (e.code !== "auth") emit("drishti:problem", e.message); return null; }
}
// Clients with their recent periods for the Overview. params: {} for the latest, {type, from, to} for one
// month or week, {since} for a download window. On this computer only, every client is read in full.
async function loadOverview(params, list) {
  if (REMOTE) {
    try { return ((await api("overview", params || {})).value || []).map((c) => ({ ...c, periods: sortPeriods(c.periods) })); }
    catch (e) { if (e.code !== "auth") emit("drishti:problem", e.message); return []; }
  }
  const out = [];
  for (const c of list || []) { if (c.archived) continue; const full = await sGet(clientKey(c.id)); if (full) out.push({ ...full, periods: sortPeriods(full.periods) }); }
  return out;
}

/* ---------------------------- formats/rules ---------------------------- */
const FMT = { PERCENT: "percent", DAYS: "days", COUNT: "count" };
const R = {
  higher: (good, fair) => ({ dir: "higher", good, fair, fmt: FMT.PERCENT }),
  lower: (good, fair, fmt = FMT.PERCENT) => ({ dir: "lower", good, fair, fmt }),
  range: (min, max) => ({ dir: "range", min, max, fmt: FMT.PERCENT }),
  count0: () => ({ dir: "lower", good: 0, fair: 0, fmt: FMT.COUNT }),
  info: (fmt = FMT.PERCENT) => ({ dir: "info", fmt }),
};

/* -------------------------- scopes of work ----------------------------- */
/* types: calc | pct | count | status | yesno | text  (kpi scopes)
          tasks (SLA / Credentialing / PCC)                              */
const SCOPES = [
  { key: "ar", name: "Account Receivable", short: "AR", kind: "kpi", group: "kpi", hasAging: true, sla: ["Production SLA"], criteria: [
    { key: "arAging", name: "AR Aging", type: "pct", std: "Less than 15% older than 90 days", formula: "Balance in >90-day buckets ÷ Total AR balance × 100", rule: R.lower(0.15, 0.15) },
    { key: "collectionRate", name: "Collection Rate", type: "calc", fmt: FMT.PERCENT, std: "Net collection rate · payments ÷ (charges − contractual adjustments)", rule: R.higher(0.85, 0.70),
      formula: "Payments ÷ (Charges − Contractual Adjustments) × 100",
      inputs: [
        { id: "payments", label: "Payments received", money: true, hint: "payment posting report" },
        { id: "charges", label: "Charges (billed)", money: true, hint: "payment posting report" },
        { id: "adjustments", label: "Contractual adjustments", money: true, hint: "payment posting report" },
      ],
      compute: (v) => { const p = n(v.payments), c = n(v.charges), a = n(v.adjustments); return (p != null && c != null && a != null && (c - a) !== 0) ? p / (c - a) : null; } },
    { key: "denialRate", name: "Denial Rate", type: "calc", fmt: FMT.PERCENT, std: "<5% excellent · 5–10% good · >10% red", formula: "Denial count ÷ Posted claim count × 100", rule: { dir: "lower", good: 0.10, fair: 0.10, best: 0.05, fmt: FMT.PERCENT },
      inputs: [{ id: "denied", label: "Denied (count)", hint: "denial report" }, { id: "submitted", label: "Posted claims (count)", hint: "payment posting report" }],
      compute: (v) => (n(v.denied) != null && n(v.submitted)) ? n(v.denied) / n(v.submitted) : null },
    { key: "daysInAR", name: "Days in AR", type: "calc", fmt: FMT.DAYS, std: "≤30 best practice · 31–40 acceptable · >40 issue", rule: R.lower(30, 40, FMT.DAYS),
      formula: "Total AR ÷ (Total payments ÷ days in the period)",
      inputs: [{ id: "totalAR", label: "Total AR inventory ($)", money: true, hint: "AR inventory report" }, { id: "avgDaily", label: "Avg daily payment ($)", money: true, helper: "avgDaily" }],
      compute: (v) => (n(v.totalAR) != null && n(v.avgDaily)) ? n(v.totalAR) / n(v.avgDaily) : null },
    { key: "appeal", name: "Appeal %", type: "pct", std: "<5% green · 5–10% amber · >10% red", formula: "Appeals filed ÷ Denials received × 100", rule: R.lower(0.05, 0.10) },
    { key: "quality", name: "Quality", type: "pct", std: "≥98% green · 95–<98% amber · <95% red", formula: "Average of Final Audit Score % (audit report)", rule: R.higher(0.98, 0.95) },
    { key: "escalations", name: "Escalations", type: "count", std: "0 green · 1+ red", formula: "Count of escalations (SR no. in tracker)", rule: R.count0() },
  ]},
  { key: "submission", name: "Submission", short: "Submission", kind: "kpi", group: "kpi", sla: ["Claim Submission", "Claim Rejection"], criteria: [
    { key: "cleanClaim", name: "Clean Claim", type: "pct", std: "90%", rule: R.higher(0.90, 0.80) },
    { key: "submissionTAT", name: "Submission TAT", type: "status", std: "24hr" },
    { key: "inflowVsSubmitted", name: "Inflow vs Submitted", type: "pct", std: "100%", rule: R.higher(1, 0.95) },
    { key: "rejectionRate", name: "Rejection Rate", type: "pct", std: "5%", rule: R.lower(0.05, 0.10) },
    { key: "resubmission", name: "Resubmission", type: "pct", std: "100%", rule: R.higher(1, 0.90) },
    { key: "quality", name: "Quality", type: "pct", std: "99%", rule: R.higher(0.99, 0.95) },
    { key: "escalations", name: "Escalations", type: "count", std: "0", rule: R.count0() },
  ]},
  { key: "ev", name: "Benefit Verification", short: "EV", kind: "kpi", group: "kpi", sla: ["Benefits Verification"], criteria: [
    { key: "completionRate", name: "Verification Completion %", type: "pct", std: "Completed ÷ Total", formula: "Completed verifications ÷ Total verifications × 100", rule: R.higher(0.95, 0.90) },
    { key: "newPatientsTAT", name: "New Patients TAT", type: "status", std: "Same Day" },
    { key: "urgentRequestTAT", name: "Urgent Request TAT", type: "status", std: "3hr" },
    { key: "tat", name: "TAT", type: "status", std: "48–72hr" },
    { key: "urgentVsRegular", name: "Urgent vs Regular verification", type: "text", std: "count split", formula: "Urgent count vs Regular count (from EV file)" },
    { key: "quality", name: "Quality", type: "pct", std: "99%", rule: R.higher(0.99, 0.95) },
    { key: "escalations", name: "Escalations", type: "count", std: "0", rule: R.count0() },
  ]},
  { key: "pp", name: "Payment Posting", short: "Payment Posting", kind: "kpi", group: "kpi", sla: ["Payment Posting"], criteria: [
    { key: "ppTAT", name: "Payment Posting TAT", type: "status", std: "24hr" },
    { key: "paymentCycle", name: "Payment Cycle", type: "calc", fmt: FMT.DAYS, std: "≤30 good · 31–45 acceptable · >45 flag", formula: "Avg (Paid date − Submission date) in days", rule: R.lower(30, 45, FMT.DAYS) },
    { key: "underpaymentRate", name: "Underpayment Rate", type: "pct", std: "<5% good, else review", formula: "Total (Allowed − Paid, where Paid<Allowed) ÷ Total Allowed × 100", rule: R.lower(0.05, 0.05) },
    { key: "varianceRate", name: "Variance Rate", type: "pct", std: "≤2% good, 2 to 5% watch, over 5% flag", formula: "Discrepancy amount (from discrepancy file) ÷ total received for posting × 100", rule: R.lower(0.02, 0.05) },
    { key: "paymentAdjustmentRate", name: "Payment Adjustment Rate", type: "pct", std: "≤5% good, 5 to 10% watch, over 10% flag", formula: "Write-off excluding CO45 ÷ gross charges billed × 100", rule: R.lower(0.05, 0.10) },
    { key: "quality", name: "Quality", type: "pct", std: "99%", rule: R.higher(0.99, 0.95) },
    { key: "topPayer", name: "Top Payer (include %)", type: "text" },
    { key: "billedVsAllowed", name: "Billed vs Allowed", type: "text" },
    { key: "websiteTrack", name: "Website Tracker maintained", type: "yesno", std: "Yes / No, set by supervisor" },
    { key: "recoupmentRate", name: "Recoupment Rate", type: "pct", std: "<5% good, else review", formula: "Amount recouped ÷ total payment received × 100", rule: R.lower(0.05, 0.05) },
    { key: "patientVsInsurance", name: "Patient vs Insurance Collection", type: "text", std: "ratio, patient ÷ insurance", formula: "Patient collection ÷ Insurance collection" },
    { key: "escalations", name: "Escalations", type: "count", std: "0", rule: R.count0() },
  ]},
  { key: "cred", name: "Credentialing", short: "Credentialing", kind: "tasks", group: "sla",
    tasks: ["Initiated Applications Vs Received Applications", "Submitted Applications", "Pending Applications (Include Client Name)", "Applications Pending With Purview", "Applications Pending With Clients", "Follow Up or Touched", "Escalations"] },
  { key: "pcc", name: "PCC", short: "PCC", kind: "tasks", group: "sla",
    criteria: [
      { key: "inflowVsPicked", name: "Production Inflow vs Picked-up Calls", type: "pct", std: "≥90% green, 80 to 89% amber, under 80% red", formula: "Picked-up calls ÷ inflow calls × 100", rule: R.higher(0.90, 0.80) },
      { key: "missedCalls", name: "Missed Calls", type: "count", std: "count" },
      { key: "callbacksMade", name: "Callbacks Made on Missed Calls", type: "count", std: "count" },
      { key: "callbackRate", name: "Callback Rate", type: "pct", std: "≥90% green, 80 to 89% amber, under 80% red", formula: "Callbacks made ÷ missed calls × 100", rule: R.higher(0.90, 0.80) },
      { key: "callsDropped", name: "Calls Dropped (in queue or other)", type: "count", std: "count" },
      { key: "dropRate", name: "Call Drop Rate", type: "pct", std: "lower is better, ≤5% green, 5 to 10% amber, over 10% red", formula: "Calls dropped ÷ total inflow calls × 100", rule: R.lower(0.05, 0.10) },
      { key: "avgHoldTime", name: "Average Hold Time", type: "text", std: "time" },
      { key: "avgHandlingTime", name: "Average Call Handling Time", type: "text", std: "time" },
      { key: "avgSpeedAnswer", name: "Average Speed of Answer", type: "text", std: "time" },
      { key: "perPersonTarget", name: "Per Person Target vs Actual Calls", type: "text", std: "user-wise" },
      { key: "outboundTargetVsActual", name: "Outbound Target vs Actual Calls", type: "pct", std: "≥90% green, 80 to 89% amber, under 80% red", formula: "Actual outbound ÷ outbound target × 100", rule: R.higher(0.90, 0.80) },
      { key: "correspondence", name: "Correspondence Received vs Sent", type: "pct", std: "≥90% green, 80 to 89% amber, under 80% red", formula: "Sent ÷ received × 100", rule: R.higher(0.90, 0.80) },
      { key: "faxEmail", name: "Fax / Email Received vs Completed", type: "pct", std: "≥90% green, 80 to 89% amber, under 80% red", formula: "Completed ÷ received × 100", rule: R.higher(0.90, 0.80) },
    ],
    tasks: ["Available Time vs Pause Time", "Online Appointment Scheduled", "Appointment Confirmation", "Missing Information Requests", "Missed Calls", "Escalations"] },
];
const SCOPE_BY_KEY = Object.fromEntries(SCOPES.map((s) => [s.key, s]));
/* Formula book: one line per KPI saying where its numbers come from. `f` supplies the formula text where the
   KPI definition did not carry one, so the info cards and the Formulas page always agree with the calculation. */
const TYPED = "Typed in by the supervisor";
const FORMULA_BOOK = {
  "ar.arAging": { from: "AR Inventory report" },
  "ar.collectionRate": { from: "Payment Posting report" },
  "ar.denialRate": { from: "Denial report and Payment Posting report" },
  "ar.daysInAR": { from: "AR Inventory report, Payment Posting report and the days in the selected duration", page: "Total AR ÷ (Total payments ÷ days in the period)" },
  "ar.appeal": { from: "AR Inventory report (comment or status column) and Denial report" },
  "ar.quality": { from: "AR Audit report" },
  "ar.escalations": { from: "Escalation Tracker" },
  "submission.cleanClaim": { from: "Claim Submission, Claim Rejection and Denial reports", f: "(Submitted − Rejected − Denied) ÷ Submitted × 100" },
  "submission.submissionTAT": { from: TYPED, manual: true, f: "Marked Met or Not met against the standard" },
  "submission.inflowVsSubmitted": { from: "Claim Submission and Submission Inflow reports", f: "Submitted claims ÷ Inflow claims × 100" },
  "submission.rejectionRate": { from: "Claim Rejection and Claim Submission reports", f: "Rejected claims ÷ Submitted claims × 100" },
  "submission.resubmission": { from: TYPED, manual: true, f: "Denied claims resubmitted ÷ Denied claims × 100" },
  "submission.quality": { from: "Submission Audit report", f: "Average of Final Audit Score % (audit report)" },
  "submission.escalations": { from: "Escalation Tracker, Submission rows", f: "Count of escalations for this department" },
  "ev.completionRate": { from: "EV / Benefit Verification report" },
  "ev.newPatientsTAT": { from: TYPED, manual: true, f: "Marked Met or Not met against the standard" },
  "ev.urgentRequestTAT": { from: TYPED, manual: true, f: "Marked Met or Not met against the standard" },
  "ev.tat": { from: TYPED, manual: true, f: "Marked Met or Not met against the standard" },
  "ev.urgentVsRegular": { from: "EV / Benefit Verification report" },
  "ev.quality": { from: TYPED, manual: true },
  "ev.escalations": { from: "Escalation Tracker, Benefit Verification rows", f: "Count of escalations for this department" },
  "pp.ppTAT": { from: TYPED, manual: true, f: "Marked Met or Not met against the standard" },
  "pp.paymentCycle": { from: "Submission file and Paid-claims file, matched claim by claim" },
  "pp.underpaymentRate": { from: "Patient-wise Ledger" },
  "pp.varianceRate": { from: "Variance / Discrepancy file and Amount Received for Posting file" },
  "pp.paymentAdjustmentRate": { from: "Non-Contractual Adjustment file and Payment Posting report" },
  "pp.quality": { from: "Payment Posting Audit report", f: "Average of Final Audit Score % (audit report)" },
  "pp.topPayer": { from: TYPED, manual: true },
  "pp.billedVsAllowed": { from: TYPED, manual: true },
  "pp.websiteTrack": { from: TYPED, manual: true, f: "Yes or No" },
  "pp.recoupmentRate": { from: "Recoupment file and Payment Posting report" },
  "pp.patientVsInsurance": { from: "Payment Posting report (payer type column)" },
  "pp.escalations": { from: "Escalation Tracker, Payment Posting rows", f: "Count of escalations for this department" },
  "pcc.inflowVsPicked": { from: "Call Log" },
  "pcc.missedCalls": { from: "Call Log", f: "Count of missed calls" },
  "pcc.callbacksMade": { from: "Callbacks on Missed Calls file", f: "Count of completed callbacks" },
  "pcc.callbackRate": { from: "Callbacks on Missed Calls file and Call Log" },
  "pcc.callsDropped": { from: "Call Log", f: "Count of dropped calls" },
  "pcc.dropRate": { from: "Call Log" },
  "pcc.avgHoldTime": { from: TYPED, manual: true },
  "pcc.avgHandlingTime": { from: TYPED, manual: true },
  "pcc.avgSpeedAnswer": { from: TYPED, manual: true },
  "pcc.perPersonTarget": { from: TYPED, manual: true },
  "pcc.outboundTargetVsActual": { from: TYPED, manual: true },
  "pcc.correspondence": { from: "Correspondence file" },
  "pcc.faxEmail": { from: "Fax / Email file" },
};
for (const sc of SCOPES) for (const c of sc.criteria || []) { const b = FORMULA_BOOK[sc.key + "." + c.key]; if (b && b.f && !c.formula) c.formula = b.f; }
// Plain-words target bands for a rule, exactly as statusFor applies them.
function ruleBands(rule) {
  if (!rule || rule.dir === "info") return [];
  const v = (x) => fmtVal(x, rule.fmt);
  if (rule.dir === "range") return [{ cls: "ok", t: `${v(rule.min)} to ${v(rule.max)}` }, { cls: "miss", t: "well outside that band" }];
  if (rule.dir === "lower" && rule.fmt === FMT.COUNT && rule.good === 0 && rule.fair === 0) return [{ cls: "ok", t: "0" }, { cls: "miss", t: "1 or more" }];
  const hi = rule.dir === "higher"; const out = [];
  if (rule.best != null) out.push({ cls: "ok", t: `Excellent: ${hi ? "at or above" : "at or below"} ${v(rule.best)}` });
  out.push({ cls: "ok", t: `${hi ? "At or above" : "At or below"} ${v(rule.good)}` });
  if (rule.fair !== rule.good) out.push({ cls: "warn", t: hi ? `${v(rule.fair)} to just under ${v(rule.good)}` : `Above ${v(rule.good)}, up to ${v(rule.fair)}` });
  out.push({ cls: "miss", t: `${hi ? "Below" : "Above"} ${v(rule.fair)}` });
  return out;
}

/* per-scope colour + short badge, for a Google-Health style colour-coded workspace */
const SCOPE_STYLE = {
  ar: { c: "#0F9F88", b: "AR" }, submission: { c: "#46B866", b: "SUB" }, ev: { c: "#33B7E0", b: "EV" },
  pp: { c: "#3E7FD0", b: "PP" }, cred: { c: "#D79432", b: "CR" }, pcc: { c: "#7B6FE0", b: "PCC" }, sla: { c: "#64748b", b: "SLA" },
};
const scopeColor = (k) => SCOPE_STYLE[k]?.c || "#2CC9B0";
const scopeBadge = (k) => SCOPE_STYLE[k]?.b || (SCOPE_BY_KEY[k]?.short || "•").slice(0, 3).toUpperCase();
const critOf = (scopeKey, critKey) => SCOPE_BY_KEY[scopeKey]?.criteria?.find((c) => c.key === critKey);

/* ------------------------------ scoring -------------------------------- */
function statusFor(actual, rule) {
  if (!isNum(actual) || !rule) return "info";
  if (rule.dir === "higher") { if (rule.best != null && actual >= rule.best) return "great"; return actual >= rule.good ? "ok" : actual >= rule.fair ? "warn" : "miss"; }
  if (rule.dir === "lower") { if (rule.best != null && actual <= rule.best) return "great"; return actual <= rule.good ? "ok" : actual <= rule.fair ? "warn" : "miss"; }
  if (rule.dir === "range") { if (actual >= rule.min && actual <= rule.max) return "ok"; const t = ((rule.max - rule.min) || 1) * 0.5; return (actual >= rule.min - t && actual <= rule.max + t) ? "warn" : "miss"; }
  return "info";
}
function critStatus(crit, cell, rule) {
  if (!cell) return "info";
  if (crit.type === "status") return cell.value === "Met" ? "ok" : cell.value === "Not Met" ? "miss" : "info";
  if (crit.type === "yesno") return cell.value === "Yes" ? "ok" : cell.value === "No" ? "miss" : "info";
  if (crit.type === "text") return "info";
  return statusFor(cell.value, rule);
}
const fmtVal = (v, fmt) => {
  if (!isNum(v)) return v == null || v === "" ? "—" : String(v);
  switch (fmt) {
    case FMT.PERCENT: return (v * 100).toFixed((v * 100) % 1 === 0 ? 0 : 1) + "%";
    case FMT.DAYS: { const x = Math.round(v); return x + (x === 1 ? " day" : " days"); }
    case FMT.COUNT: return Math.round(v).toLocaleString();
    default: return v.toLocaleString(undefined, { maximumFractionDigits: 2 });
  }
};
// how a metric's number is shown: its own format, its rule's format, a plain count, or a percentage
const critFmt = (crit) => crit.fmt || crit.rule?.fmt || (crit.type === "count" ? FMT.COUNT : FMT.PERCENT);
function critDisplay(crit, cell) {
  if (!cell) return "—";
  if (["status", "yesno", "text"].includes(crit.type)) return cell.value || "—";
  return fmtVal(cell.value, critFmt(crit));
}
const targetText = (r) => !r ? "" :
  r.dir === "higher" ? "≥ " + fmtVal(r.good, r.fmt) : r.dir === "lower" ? "≤ " + fmtVal(r.good, r.fmt) :
  r.dir === "range" ? fmtVal(r.min, r.fmt) + "–" + fmtVal(r.max, r.fmt) : "";
const stLabel = (s) => (s === "great" ? "Excellent" : s === "ok" ? "On target" : s === "warn" ? "At risk" : s === "miss" ? "Missed" : "—");

/* --------------------- generic report reader --------------------------- */
async function readReport(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  let hi = 0;
  for (let i = 0; i < Math.min(rows.length, 15); i++) { if ((rows[i] || []).filter((c) => typeof c === "string" && c.trim()).length >= 2) { hi = i; break; } }
  const headers = (rows[hi] || []).map((h, i) => (h == null || h === "" ? `Column ${i + 1}` : String(h).trim()));
  const dataRows = rows.slice(hi + 1).filter((r) => (r || []).some((c) => c != null && c !== ""));
  const columns = headers.map((name, ci) => { let sum = 0, numeric = 0; for (const r of dataRows) { const c = (r || [])[ci]; if (c != null && c !== "") { const x = n(c); if (x != null) { sum += x; numeric++; } } } return { name, index: ci, sum, numeric }; });
  return { rowCount: dataRows.length, columns, headers, rows: dataRows, fileName: file.name };
}

/* ---- report metadata (card display) ---- */
const REPORT_DEFS = [
  { key: "ar", name: "AR Inventory report (current period)", tint: "#eaf4ff", hint: "Balance, aging (or DOS), status and comments for appeals" },
  { key: "arprev", name: "AR Inventory report (previous period)", tint: "#e3eefb", hint: "For the aging comparison: balance + aging (or DOS)" },
  { key: "denial", name: "Denial report", tint: "#fdeeee", hint: "Denied claim count" },
  { key: "payment", name: "Payment Posting report", tint: "#e7f6f3", hint: "Payments, charges, contractual adjustments, claim count" },
  { key: "submission", name: "Claim Submission report", tint: "#eef7e9", hint: "Submitted claim count" },
  { key: "rejection", name: "Claim Rejection report", tint: "#fdeeee", hint: "Rejected claim count" },
  { key: "inflow", name: "Submission Inflow report", tint: "#eef4ff", hint: "Claims received (from PMS/client)" },
  { key: "audit", name: "AR Audit report", tint: "#fef6e3", hint: "Final Audit Score %" },
  { key: "subaudit", name: "Submission Audit report", tint: "#fef6e3", hint: "Final Audit Score %" },
  { key: "ppaudit", name: "Payment Posting Audit report", tint: "#fef6e3", hint: "Final Audit Score %" },
  { key: "escalation", name: "Escalation Tracker", tint: "#f3eefe", hint: "Escalation count (SR no.)" },
  { key: "ev", name: "EV / Benefit Verification report", tint: "#f3eefe", hint: "Completed / Not-completed + Urgent / Regular columns" },
  { key: "paysub", name: "Submission file (for Payment Cycle)", tint: "#eef4ff", hint: "Claim identifier + Submission date" },
  { key: "paypaid", name: "Paid-claims file (for Payment Cycle)", tint: "#e7f6f3", hint: "Claim identifier + Paid date" },
  { key: "ledger", name: "Patient-wise Ledger (Payment Posting)", tint: "#fef6e3", hint: "Allowed + Paid amounts (for Underpayment Rate)" },
  { key: "postrecv", name: "Amount Received for Posting file (Payment Posting)", tint: "#eef4ff", hint: "All accounts that came in for posting, the denominator for Variance Rate" },
  { key: "variance", name: "Variance / Discrepancy file (Payment Posting)", tint: "#fdeeee", hint: "Amount that could not be posted. This is the source used for Variance Rate" },
  { key: "recoup", name: "Recoupment file (Payment Posting)", tint: "#fdeeee", hint: "Amount recouped or taken back by payers (for Recoupment Rate)" },
  { key: "noncontr", name: "Non-Contractual Adjustment file (Payment Posting)", tint: "#f3eefe", hint: "Write-offs excluding CO45, e.g. TFL (for Payment Adjustment Rate)" },
  { key: "cred", name: "Credentialing report", tint: "#eef2f6", hint: "Applications + status" },
  { key: "pcccalls", name: "Call Log (PCC)", tint: "#eef4ff", hint: "One row per inbound call with a disposition/status column (picked, missed, dropped)" },
  { key: "pcccallback", name: "Callbacks on Missed Calls (PCC)", tint: "#e7f6f3", hint: "Callback attempts with a completed / not-completed column" },
  { key: "pcccorr", name: "Correspondence (PCC)", tint: "#f3eefe", hint: "Correspondence received, with a completed / not-completed (sent) column" },
  { key: "pccfax", name: "Fax / Email (PCC)", tint: "#fef6e3", hint: "Fax / email received, with a completed / not-completed column" },
];
const REPORT_BY_KEY = Object.fromEntries(REPORT_DEFS.map((d) => [d.key, d]));
/* which report upload slots each scope of work needs */
const SCOPE_REPORTS = { ar: ["ar", "denial", "audit", "escalation"], submission: ["submission", "rejection", "inflow", "subaudit"], pp: ["payment", "paysub", "paypaid", "ledger", "postrecv", "variance", "noncontr", "recoup", "ppaudit"], ev: ["ev"], cred: [], pcc: ["pcccalls", "pcccallback", "pcccorr", "pccfax"], sla: [] };
/* the upload slots a scope shows for a given client. Account Receivable needs the Payment Posting report for
   Collection Rate, Denial Rate and Days in AR, so it is shown there when the client has no Payment Posting scope. */
const scopeReportKeys = (scopeKey, clientScopes) => { const base = SCOPE_REPORTS[scopeKey] || []; return scopeKey === "ar" && !(clientScopes || []).includes("pp") ? [...base.slice(0, 2), "payment", ...base.slice(2)] : base; };
/* reports that feed SLA breakdowns only (kept separate so KPI uploads stay simple) */
const SCOPE_SLA_REPORTS = { ar: ["arprev"], submission: [], pp: [], ev: [], cred: [], pcc: [], sla: [] };

/* fields to detect per report; agg: sum | avg | col. optional fields never block. */
const P = {
  balance: [/ar\s*balance/i, /remaining\s*balance/i, /outstanding/i, /\bbalance\b/i, /amount\s*due/i, /open\s*amount/i],
  charge: [/total\s*charge/i, /charge\s*amount/i, /charges/i, /\bcharge\b/i, /billed\s*amount/i, /billed/i],
  payment: [/payment\s*amount/i, /amount\s*paid/i, /paid\s*amount/i, /net\s*paid/i, /\bpayment\b/i, /\bpaid\b/i],
  adj: [/contractual\s*adj/i, /contractual\s*write/i, /contractual/i, /write\s*-?off/i, /adjustment\s*amount/i, /adjustment/i],
  dos: [/date\s*of\s*service/i, /\bdos\b/i, /service\s*date/i, /svc\s*date/i],
  bucket: [/aging\s*bucket/i, /\baging\b/i, /\bbucket\b/i, /age\s*category/i],
  score: [/final\s*audit\s*score/i, /audit\s*score/i, /final\s*score/i, /quality\s*score/i, /\bscore\b/i],
  appeal: [/comment/i, /sub\s*status/i, /\bstatus\b/i, /\baction\b/i, /notes/i],
  dept: [/department/i, /\bdept\b/i, /\bscope\b/i, /work\s*of\s*scope/i],
  claimid: [/claim\s*id/i, /claim\s*no/i, /claim\s*#/i, /claim\s*number/i, /claim\s*num/i, /\bclaim\b/i, /icn/i],
  action: [/\baction\b/i, /sub\s*status/i, /current\s*status/i, /\bstatus\b/i, /disposition/i],
  // any unique identifier usable to join two files (broadest match-key detector)
  idany: [/claim\s*id/i, /claim\s*no/i, /claim\s*#/i, /claim\s*number/i, /claim\s*num/i, /\bicn\b/i, /\bdcn\b/i, /account\s*(no|number|#)/i, /\bacct\b/i, /encounter\s*(no|number|#|id)/i, /\bencounter\b/i, /reference\s*(no|number|#)/i, /\bref\s*(no|#)/i, /\bmrn\b/i, /patient\s*(id|account)/i],
  // stable fields for a composite key (do NOT use balance across lifecycle stages)
  patient: [/patient\s*name/i, /\bpatient\b/i, /member\s*name/i, /subscriber/i, /\bname\b/i],
  cpt: [/\bcpt\b/i, /procedure\s*code/i, /\bproc\b/i, /hcpcs/i, /service\s*code/i],
  payer: [/\bpayer\b/i, /\bpayor\b/i, /insurance/i, /\bcarrier\b/i, /\bplan\b/i],
  provider: [/provider/i, /\brendering\b/i, /physician/i, /\bdoctor\b/i, /\bnpi\b/i],
  billed: [/billed\s*amount/i, /\bbilled\b/i, /total\s*charge/i, /charge\s*amount/i, /\bcharges?\b/i],
  // new-metric fields
  allowed: [/allowed\s*amount/i, /\ballowed\b/i, /allowable/i, /contract(ed)?\s*rate/i, /fee\s*schedule/i],
  discrepancy: [/discrepanc/i, /variance/i, /un-?posted/i, /unable\s*to\s*post/i, /not\s*posted/i, /difference/i, /\bamount\b/i, /\bbalance\b/i],
  writeoff: [/write.?off/i, /\bw\/?o\b/i, /adjustment\s*amount/i, /adjustment/i, /\bamount\b/i],
  postingin: [/received\s*for\s*posting/i, /amount\s*received/i, /posting\s*amount/i, /remit(tance)?\s*amount/i, /era\s*amount/i, /payment\s*amount/i, /amount\s*paid/i, /\bpaid\b/i],
  code: [/reason\s*code/i, /adj(ustment)?\s*code/i, /\bcarc\b/i, /denial\s*code/i, /\bcode\b/i, /reason/i],
  recoup: [/recoup/i, /recovery/i, /take-?back/i, /reversal/i, /refund/i, /\bamount\b/i, /\bbalance\b/i],
  source: [/payer\s*type/i, /payment\s*source/i, /paid\s*by/i, /responsibility/i, /payment\s*from/i, /\bsource\b/i, /\btype\b/i],
  subdate: [/submission\s*date/i, /submitted\s*date/i, /date\s*submitted/i, /\bbill(ed)?\s*date/i, /claim\s*date/i, /\bsub\s*date/i],
  paiddate: [/paid\s*date/i, /payment\s*date/i, /date\s*paid/i, /check\s*date/i, /remit(tance)?\s*date/i, /post(ed|ing)?\s*date/i],
  completed: [/completed?/i, /\bstatus\b/i, /verification\s*status/i, /complete\s*flag/i, /done/i, /sent/i, /disposition/i],
  urgency: [/urgent/i, /priority/i, /request\s*type/i, /verification\s*type/i, /\btype\b/i],
  // PCC call log disposition (picked / missed / dropped)
  disposition: [/disposition/i, /call\s*status/i, /call\s*result/i, /\bstatus\b/i, /\bresult\b/i, /outcome/i],
};
const DEPT_PATS = { ar: /receivab|\bar\b|a\/?r\b/i, submission: /submiss|billing|submit|claims?\s*sub/i, pp: /posting|payment/i, ev: /benefit|verif|eligib|\bev\b/i, cred: /credential/i, pcc: /\bpcc\b|patient\s*care/i };
const FIELDS = {
  ar: [{ id: "balance", label: "Balance amount", agg: "sum", pats: P.balance }],
  arprev: [{ id: "balance", label: "Balance amount", agg: "sum", pats: P.balance }],
  payment: [{ id: "payments", label: "Payments received", agg: "sum", pats: P.payment }, { id: "charges", label: "Charges (billed)", agg: "sum", pats: P.charge }, { id: "adjustments", label: "Contractual adjustments", agg: "sum", pats: P.adj }],
  audit: [{ id: "score", label: "Final Audit Score %", agg: "avg", pats: P.score }],
  subaudit: [{ id: "score", label: "Final Audit Score %", agg: "avg", pats: P.score }],
  ppaudit: [{ id: "score", label: "Final Audit Score %", agg: "avg", pats: P.score }],
  denial: [], submission: [], escalation: [], ev: [], cred: [], pcc: [],
  paysub: [], paypaid: [], ledger: [], postrecv: [], variance: [], noncontr: [], recoup: [],
  pcccalls: [], pcccallback: [], pcccorr: [], pccfax: [],
};

/* ---------- choosing the right column ----------
   A column is used only when Drishti is sure. Three checks, in order:
     1. the heading must match what we are looking for,
     2. the column must hold the right kind of data (an amount column must hold numbers, and a date,
        ID, code, name or status column is never used as an amount),
     3. exactly one column must be a clear match.
   If no column passes, or more than one is a clear match, nothing is guessed: the supervisor is asked. */
const GENERIC_WORDS = new Set(["amount", "balance", "type", "status", "paid", "payment", "claim", "name", "code", "reason", "score", "result", "adjustment", "charge", "charges", "charges?", "billed", "difference", "done", "sent", "action", "notes", "comment", "aging", "bucket", "plan", "insurance", "source", "priority", "outcome", "allowed", "refund", "reversal", "recovery", "variance", "scope", "proc", "provider", "patient"]);
const isGenericPat = (p) => GENERIC_WORDS.has(p.source.replace(/\\b/g, "").toLowerCase());
const NOT_AMOUNT_HDR = /\bdate\b|\bdob\b|\bdos\b|\btime\b|\bid\b|\bcode\b|\bno\b|\bnum(ber)?\b|#|\bcpt\b|\bname\b|\btype\b|\bstatus\b|\bphone\b|\bzip\b|\bnpi\b/i;
const NOT_FLAG_HDR = /\bdate\b|\btime\b|\bamount\b|\bid\b|\bno\b|#/i;
/* what each looked-up item is: ask=true means a KPI depends on it, so Drishti must be sure or ask */
const FIELD_META = {
  balance: { kind: "amount", ask: true, label: "Balance amount" },
  charges: { kind: "amount", ask: true, label: "Charges (billed)" },
  payments: { kind: "amount", ask: true, label: "Payments received" },
  adjustments: { kind: "amount", ask: true, label: "Contractual adjustments" },
  score: { kind: "amount", ask: true, label: "Final Audit Score %" },
  recoup: { kind: "amount", ask: true, label: "Recouped amount" },
  allowed: { kind: "amount", ask: true, label: "Allowed amount" },
  paid: { kind: "amount", ask: true, label: "Paid amount" },
  received: { kind: "amount", ask: true, label: "Amount received for posting" },
  discrepancy: { kind: "amount", ask: true, label: "Discrepancy amount" },
  writeoff: { kind: "amount", ask: true, label: "Write-off / adjustment amount" },
  date: { kind: "date", ask: true, label: "Date" },
  completed: { kind: "flag", ask: true, label: "Completed / Not-completed" },
  urgency: { kind: "flag", ask: true, label: "Urgent / Regular" },
  disposition: { kind: "flag", ask: true, label: "Call disposition / status" },
  paidcol: { kind: "amount" }, grpval: { kind: "amount" }, archarges: { kind: "amount" },
};
const hdrText = (h) => String(h == null ? "" : h).replace(/_+/g, " ").replace(/\s+/g, " ").trim();
const colStats = new WeakMap();
function nonEmptyCount(rep, idx) {
  let m = colStats.get(rep); if (!m) { m = {}; colStats.set(rep, m); }
  if (m[idx] == null) { let c = 0; for (const r of rep.rows) { const v = (r || [])[idx]; if (v != null && v !== "") c++; } m[idx] = c; }
  return m[idx];
}
function colFits(rep, idx, kind) {
  const h = hdrText(rep.headers[idx]);
  if (kind === "amount") {
    if (NOT_AMOUNT_HDR.test(h)) return false;
    const filled = nonEmptyCount(rep, idx); if (!filled) return true;           // an empty column cannot be judged, the heading decides
    const numeric = rep.columns && rep.columns[idx] ? rep.columns[idx].numeric : 0;
    return numeric >= filled * 0.6;
  }
  if (kind === "flag") return !NOT_FLAG_HDR.test(h);
  return true;
}
// Every column that could be this item, best first, and whether one of them is a clear winner.
function resolveCol(rep, pats, kind) {
  const c = [];
  rep.headers.forEach((h0, i) => {
    const h = hdrText(h0);
    let rank = -1, distinct = false;
    for (let k = 0; k < pats.length; k++) if (pats[k].test(h)) { if (rank < 0) rank = k; if (!isGenericPat(pats[k])) { distinct = true; rank = Math.min(rank, k); break; } }
    if (rank < 0 || !colFits(rep, i, kind)) return;
    c.push({ i, rank, distinct });
  });
  if (!c.length) return { idx: null, options: [], sure: false };
  c.sort((a, b) => (b.distinct - a.distinct) || (a.rank - b.rank) || (a.i - b.i));
  const distinct = c.filter((x) => x.distinct);
  const sure = distinct.length === 1 || (distinct.length === 0 && c.length === 1);
  return { idx: c[0].i, options: c.map((x) => x.i), sure };
}
let DERIVE_CTX = null;   // set while one report is being read: which columns were used, which need asking
const findIdx = (rep, pats, override, id) => {
  const meta = FIELD_META[id] || {};
  if (override && Object.prototype.hasOwnProperty.call(override, id)) {
    const v = override[id]; if (v === "" || v == null) return null;            // the supervisor said the file does not have it
    const idx = Number(v); if (DERIVE_CTX && meta.ask) { DERIVE_CTX.used[id] = idx; DERIVE_CTX.byPats.set(pats, idx); } return idx;
  }
  if (!meta.ask && DERIVE_CTX && DERIVE_CTX.byPats.has(pats)) return DERIVE_CTX.byPats.get(pats);   // reuse the column already settled for this item
  const r = resolveCol(rep, pats, meta.kind || "any");
  if (r.idx == null) return null;
  if (meta.ask && !r.sure && DERIVE_CTX) { if (!DERIVE_CTX.amb[id]) DERIVE_CTX.amb[id] = r.options; return null; }
  if (DERIVE_CTX && meta.ask) { DERIVE_CTX.used[id] = r.idx; DERIVE_CTX.byPats.set(pats, r.idx); }
  return r.idx;
};
const sumIdx = (rep, idx) => { let s = 0, f = false; for (const r of rep.rows) { const x = n((r || [])[idx]); if (x != null) { s += x; f = true; } } return f ? s : null; };
const avgIdx = (rep, idx) => { let s = 0, c = 0; for (const r of rep.rows) { const x = n((r || [])[idx]); if (x != null) { s += x; c++; } } return c ? s / c : null; };
function parseDate(v) { if (v == null || v === "") return null; if (v instanceof Date) return isNaN(v) ? null : v; if (typeof v === "number") { const d = new Date(Math.round((v - 25569) * 86400000)); return isNaN(d) ? null : d; } const d = new Date(v); return isNaN(d) ? null : d; }
function ageBucket(days) { return days <= 30 ? "0-30" : days <= 60 ? "31-60" : days <= 90 ? "61-90" : days <= 120 ? "91-120" : days <= 180 ? "121-180" : days <= 360 ? "181-360" : "360+"; }
const AR_BUCKETS = ["0-30", "31-60", "61-90", "91-120", "121-180", "181-360", "360+"];
function canonBucket(label) { if (/360/.test(label) && /\+|plus/i.test(label)) return "360+"; const nums = String(label).match(/\d+/g); const lo = nums ? parseInt(nums[0]) : 0; if (lo >= 360) return "360+"; if (lo >= 181) return "181-360"; if (lo >= 121) return "121-180"; if (lo >= 91) return "91-120"; if (lo >= 61) return "61-90"; if (lo >= 31) return "31-60"; return "0-30"; }
function countClaims(rep, override) { const idx = findIdx(rep, P.claimid, override, "claimid"); if (idx == null) return rep.rowCount; const seen = new Set(); for (const r of rep.rows) { const v = String((r || [])[idx] ?? "").trim(); if (v) seen.add(v); } return seen.size || rep.rowCount; }
function groupSum(rep, override, keyPats, valPats, keyId) {
  const kI = findIdx(rep, keyPats, override, keyId || "grpkey"); if (kI == null) return null;
  const vI = valPats ? findIdx(rep, valPats, override, "grpval") : null;
  const g = {};
  for (const r of rep.rows) { const k = String((r || [])[kI] ?? "").trim() || "(blank)"; const v = vI != null ? (n((r || [])[vI]) || 0) : 1; g[k] = (g[k] || 0) + v; }
  return Object.keys(g).length ? g : null;
}
function is90plus(bucket) { const nums = String(bucket).match(/\d+/g); const lo = nums ? parseInt(nums[0]) : null; if (lo != null && lo >= 91) return true; if (/\+/.test(bucket) && lo != null && lo >= 90) return true; return false; }
/* completed / not-completed classifier — returns true (done), false (not done), null (blank, skip) */
function isDone(v) {
  const t = String(v ?? "").trim().toLowerCase(); if (!t) return null;
  if (/\bnot\b|\bno\b|pending|incomplete|\bopen\b|fail|missed|unanswered|awaiting|\bwip\b|in\s*progress/.test(t)) return false;
  if (/complet|done|sent|reach|connect|success|resolv|answered|closed|^y$|^yes$|^1$|^true$/.test(t)) return true;
  return false;
}

/* derive one report into figures; returns { ...figures, missing:[{id,label}], columns:[...] }.
   override maps fieldId -> column index chosen by the supervisor. */
function deriveReport(key, rep, override) {
  const ctx = { used: {}, amb: {}, byPats: new Map() };
  DERIVE_CTX = ctx;
  try { return deriveReportInner(key, rep, override, ctx); } finally { DERIVE_CTX = null; }
}
const fieldLabel = (key, id) => (id === "date" ? (key === "paysub" ? "Submission date" : "Paid date") : (FIELD_META[id] && FIELD_META[id].label) || id);
function deriveReportInner(key, rep, override, ctx) {
  const out = { fileName: rep.fileName, rows: rep.rowCount, missing: [], columns: rep.columns.map((c) => ({ name: c.name, index: c.index })) };
  for (const f of (FIELDS[key] || [])) {
    const idx = findIdx(rep, f.pats, override, f.id);
    if (idx == null) { out.missing.push({ id: f.id, label: f.label }); continue; }
    out[f.id] = f.agg === "avg" ? avgIdx(rep, idx) : sumIdx(rep, idx);
    if (out[f.id] == null) out.missing.push({ id: f.id, label: f.label });
  }
  if (key === "payment") {
    out.claims = countClaims(rep, override); out.payerPay = groupSum(rep, override, P.payer, P.payment, "payer");
    // patient vs insurance collection split (needs a source/type column)
    const sI = findIdx(rep, P.source, override, "source"); const pI = findIdx(rep, P.payment, override, "paidcol");
    if (sI != null && pI != null) { let pt = 0, ins = 0; for (const r of rep.rows) { const src = String((r || [])[sI] ?? "").toLowerCase(); const amt = n((r || [])[pI]) || 0; if (/patient|\bself\b|\bpt\b|guarantor|personal/.test(src)) pt += amt; else ins += amt; } out.patientPaid = pt; out.insurancePaid = ins; }
  }
  // recoupment file → amount recouped by payers
  if (key === "recoup") {
    const rI = findIdx(rep, P.recoup, override, "recoup");
    if (rI == null) out.missing.push({ id: "recoup", label: "Recouped amount column" });
    else { out.recoupAmt = sumIdx(rep, rI); out.count = rep.rowCount; }
  }
  if (key === "denial") { out.count = countClaims(rep, override); out.payerDenialCnt = groupSum(rep, override, P.payer, null, "payer"); out.payerDenialAmt = groupSum(rep, override, P.payer, [/denied\s*amount/i, /denial\s*amount/i, /\bamount\b/i, /\bbalance\b/i, /charge/i], "payer"); }
  if (key === "escalation") {
    out.count = rep.rowCount;
    const di = findIdx(rep, P.dept, override, "dept");
    if (di != null) { const g = {}; for (const r of rep.rows) { const d = String((r || [])[di] ?? "").trim(); if (d) g[d] = (g[d] || 0) + 1; } if (Object.keys(g).length) out.byDept = g; }
  }
  if (key === "submission") out.count = countClaims(rep, override);
  if (key === "rejection") out.count = countClaims(rep, override);
  if (key === "inflow") out.count = countClaims(rep, override);
  // files kept raw for cross-file matching (Payment Cycle)
  if (key === "paysub" || key === "paypaid") {
    out.raw = { headers: rep.headers, rows: rep.rows };
    const dateP = key === "paysub" ? P.subdate : P.paiddate;
    const di = findIdx(rep, dateP, override, "date");
    if (di == null) out.missing.push({ id: "date", label: key === "paysub" ? "Submission date column" : "Paid date column" });
    else out.dateIdx = di;
    out.count = rep.rowCount;
  }
  // patient-wise ledger → underpayment (allowed vs paid, per row)
  if (key === "ledger") {
    const aI = findIdx(rep, P.allowed, override, "allowed");
    const pI = findIdx(rep, P.payment, override, "paid");
    if (aI == null) out.missing.push({ id: "allowed", label: "Allowed amount column" });
    if (pI == null) out.missing.push({ id: "paid", label: "Paid amount column" });
    if (aI != null && pI != null) {
      let totAllowed = 0, under = 0, variance = 0, rows = 0;
      for (const r of rep.rows) { const a = n((r || [])[aI]), p = n((r || [])[pI]); if (a == null) continue; totAllowed += a; rows++; if (p != null) { if (p < a) under += (a - p); variance += Math.abs(a - p); } }
      out.ledgerAllowed = totAllowed; out.ledgerUnder = under; out.ledgerVariance = variance; out.count = rows;
    }
  }
  // amount received for posting → total that came in for posting (denominator for Variance Rate)
  if (key === "postrecv") {
    const rI = findIdx(rep, P.postingin, override, "received");
    if (rI == null) out.missing.push({ id: "received", label: "Amount received for posting column" });
    else { out.postingIn = sumIdx(rep, rI); out.count = rep.rowCount; }
  }
  // variance / discrepancy file → amount that could not be posted
  if (key === "variance") {
    const dI = findIdx(rep, P.discrepancy, override, "discrepancy");
    if (dI == null) out.missing.push({ id: "discrepancy", label: "Discrepancy amount column" });
    else { out.varianceAmt = sumIdx(rep, dI); out.count = rep.rowCount; }
  }
  // non-contractual adjustment file → write-offs excluding CO45
  if (key === "noncontr") {
    const wI = findIdx(rep, P.writeoff, override, "writeoff");
    if (wI == null) out.missing.push({ id: "writeoff", label: "Write-off / adjustment amount column" });
    else {
      const cI = findIdx(rep, P.code, override, "code");
      let wo = 0, rows = 0;
      for (const r of rep.rows) { const w = n((r || [])[wI]); if (w == null) continue; if (cI != null) { const code = String((r || [])[cI] ?? "").toLowerCase(); if (/co-?\s*45|\b45\b/.test(code)) continue; } wo += w; rows++; }
      out.writeoff = wo; out.count = rows;
    }
  }
  // EV — completion + urgent/regular from two flag columns
  if (key === "ev") {
    const cI = findIdx(rep, P.completed, override, "completed");
    const uI = findIdx(rep, P.urgency, override, "urgency");
    if (cI == null) out.missing.push({ id: "completed", label: "Completed / Not-completed column" });
    if (uI == null) out.missing.push({ id: "urgency", label: "Urgent / Regular column" });
    if (cI != null) { let comp = 0, tot = 0; for (const r of rep.rows) { const v = String((r || [])[cI] ?? "").trim(); if (!v) continue; tot++; if (/^(y|yes|complete|completed|done|verified|1|true)$/i.test(v)) comp++; } out.evCompleted = comp; out.evTotal = tot; }
    if (uI != null) { let urg = 0, reg = 0; for (const r of rep.rows) { const v = String((r || [])[uI] ?? "").trim(); if (!v) continue; if (/urgent|stat|priority|rush/i.test(v)) urg++; else reg++; } out.evUrgent = urg; out.evRegular = reg; }
  }
  // PCC — call log disposition (picked / missed / dropped)
  if (key === "pcccalls") {
    const dI = findIdx(rep, P.disposition, override, "disposition");
    if (dI == null) out.missing.push({ id: "disposition", label: "Call disposition / status column (picked, missed, dropped)" });
    else {
      let picked = 0, missed = 0, dropped = 0, other = 0, total = 0;
      for (const r of rep.rows) {
        const t = String((r || [])[dI] ?? "").trim().toLowerCase(); if (!t) continue; total++;
        if (/drop|abandon|queue|disconnect|hung\s*up|hungup/.test(t)) dropped++;
        else if (/miss|no\s*answer|unanswered|not\s*answered|\bnca\b|no\s*response|voicemail|\bvm\b/.test(t)) missed++;
        else if (/pick|answer|connect|handl|complet|attend|talk|reach|receiv|inbound|\bdone\b/.test(t)) picked++;
        else other++;
      }
      out.callsTotal = total; out.callsPicked = picked; out.callsMissed = missed; out.callsDropped = dropped; out.callsOther = other;
    }
  }
  // PCC — callbacks on missed calls (completed / not-completed)
  if (key === "pcccallback") {
    const cI = findIdx(rep, P.completed, override, "completed");
    if (cI == null) out.missing.push({ id: "completed", label: "Callback completed / not-completed column" });
    else { let done = 0, tot = 0; for (const r of rep.rows) { const d = isDone((r || [])[cI]); if (d == null) continue; tot++; if (d) done++; } out.cbDone = done; out.cbTotal = tot; }
  }
  // PCC — correspondence received vs sent (completed / not-completed)
  if (key === "pcccorr") {
    const cI = findIdx(rep, P.completed, override, "completed");
    if (cI == null) out.missing.push({ id: "completed", label: "Correspondence completed / sent column" });
    else { let done = 0, tot = 0; for (const r of rep.rows) { const d = isDone((r || [])[cI]); if (d == null) continue; tot++; if (d) done++; } out.corrDone = done; out.corrTotal = tot; }
  }
  // PCC — fax / email received vs completed (completed / not-completed)
  if (key === "pccfax") {
    const cI = findIdx(rep, P.completed, override, "completed");
    if (cI == null) out.missing.push({ id: "completed", label: "Fax / email completed column" });
    else { let done = 0, tot = 0; for (const r of rep.rows) { const d = isDone((r || [])[cI]); if (d == null) continue; tot++; if (d) done++; } out.faxDone = done; out.faxTotal = tot; }
  }
  if (key === "ar" || key === "arprev") {
    out.totalAR = out.balance; { const cI = findIdx(rep, P.charge, override, "archarges"); out.totalCharges = cI != null ? sumIdx(rep, cI) : null; } out.count = countClaims(rep, override);
    if (key === "ar") out.payerAR = groupSum(rep, override, P.payer, P.balance, "payer");
    // aging: prefer bucket column, else compute from DOS (age = today - DOS)
    const bIdx = findIdx(rep, P.bucket, override, "bucket");
    const dIdx = findIdx(rep, P.dos, override, "dos");
    const balIdx = findIdx(rep, P.balance, override, "balance");
    if (balIdx != null && (bIdx != null || dIdx != null)) {
      const g = {}, gc = {}, seen = {}; const today = new Date(); const cIdx = findIdx(rep, P.claimid, override, "claimid");
      for (const r of rep.rows) {
        const bal = n((r || [])[balIdx]); if (bal == null) continue;
        let bucket = null;
        if (bIdx != null) bucket = String((r || [])[bIdx] ?? "").trim();
        if (!bucket && dIdx != null) { const d = parseDate((r || [])[dIdx]); if (d) bucket = ageBucket(Math.floor((today - d) / 86400000)); }
        if (bucket) {
          g[bucket] = (g[bucket] || 0) + bal;
          if (cIdx != null) { const id = String((r || [])[cIdx] ?? "").trim(); if (id) { seen[bucket] = seen[bucket] || new Set(); if (!seen[bucket].has(id)) { seen[bucket].add(id); gc[bucket] = (gc[bucket] || 0) + 1; } } else gc[bucket] = (gc[bucket] || 0) + 1; }
          else gc[bucket] = (gc[bucket] || 0) + 1;
        }
      }
      if (Object.keys(g).length) { out.aging = g; out.agingCount = gc; out.agingSource = bIdx != null ? "report bucket" : "computed from DOS"; }
    } else if (out.totalAR != null) {
      out.missing.push({ id: "bucket", label: "Aging bucket (or Date of Service)" });
    }
    // appeals: rows where comment/sub-status/action mention appeal submitted / in process (not rejected/denied), counted once
    const appealCols = key === "ar" ? rep.headers.map((h, i) => (P.appeal.some((p) => p.test(hdrText(h))) ? i : -1)).filter((i) => i >= 0) : [];
    if (appealCols.length) {
      let ct = 0;
      for (const r of rep.rows) {
        let hit = false;
        for (const ci of appealCols) {
          const t = String((r || [])[ci] ?? "").toLowerCase();
          if (/appeal/.test(t) && /(submit|in\s*process|in-process|inprocess)/.test(t) && !/(reject|denied)/.test(t)) { hit = true; break; }
        }
        if (hit) ct++;
      }
      out.appeals = ct;
    }
    // status-wise breakdown from the Action column (current AR only) — mandatory for the SLA Status breakdown
    if (key === "ar") {
      const actIdx = findIdx(rep, P.action, override, "action");
      const balIdx2 = findIdx(rep, P.balance, override, "balance");
      if (actIdx != null && balIdx2 != null) {
        const sb = {}, seen = {}, mat = {}; const cIdx = findIdx(rep, P.claimid, override, "claimid");
        const bkIdx = findIdx(rep, P.bucket, override, "bucket"), dsIdx = findIdx(rep, P.dos, override, "dos"); const today2 = new Date();
        for (const r of rep.rows) {
          const st = (String((r || [])[actIdx] ?? "").trim()) || "(blank)";
          const bal = n((r || [])[balIdx2]) || 0;
          sb[st] = sb[st] || { count: 0, amount: 0 }; sb[st].amount += bal;
          if (cIdx != null) { const id = String((r || [])[cIdx] ?? "").trim(); if (id) { seen[st] = seen[st] || new Set(); if (!seen[st].has(id)) { seen[st].add(id); sb[st].count++; } } else sb[st].count++; }
          else sb[st].count++;
          // status x aging bucket amount
          let bk = null;
          if (bkIdx != null) bk = String((r || [])[bkIdx] ?? "").trim();
          if (!bk && dsIdx != null) { const d = parseDate((r || [])[dsIdx]); if (d) bk = ageBucket(Math.floor((today2 - d) / 86400000)); }
          if (bk) { const cb = canonBucket(bk); mat[st] = mat[st] || {}; mat[st][cb] = (mat[st][cb] || 0) + bal; }
        }
        out.statusBreakdown = sb; out.statusAging = mat;
      } else if (actIdx == null) {
        out.missing.push({ id: "action", label: "Action column, required for the SLA Status Wise Breakdown" });
      }
    }
  }
  // columns more than one of which could fit: never guessed, always asked
  out.ambiguous = Object.entries(ctx.amb).filter(([id]) => ctx.used[id] == null).map(([id, options]) => ({ id, label: fieldLabel(key, id), options }));
  const ambIds = new Set(out.ambiguous.map((a) => a.id));
  out.missing = out.missing.filter((f, i, arr) => !ambIds.has(f.id) && arr.findIndex((x) => x.id === f.id) === i);
  // which heading each figure was read from, shown on the report card
  out.cols = Object.entries(ctx.used).map(([id, idx]) => ({ id, label: fieldLabel(key, id), name: rep.headers[idx] }));
  return out;
}

/* ===== universal match engine: join two raw files by a shared ID, else a composite key ===== */
const COMPOSITE_FIELDS = [
  { id: "patient", label: "Patient", pats: P.patient },
  { id: "dos", label: "Date of Service", pats: P.dos },
  { id: "cpt", label: "CPT", pats: P.cpt },
  { id: "payer", label: "Payer", pats: P.payer },
  { id: "provider", label: "Provider", pats: P.provider },
  { id: "billed", label: "Billed amount", pats: P.billed },
];
const idxOf = (headers, pats) => { const i = headers.findIndex((h) => pats.some((p) => p.test(hdrText(h)))); return i >= 0 ? i : null; };
function normCell(v, id) {
  if (v == null) return "";
  if (id === "dos") { const d = parseDate(v); return d ? d.toISOString().slice(0, 10) : String(v).trim().toLowerCase(); }
  if (id === "billed") { const x = n(v); return x == null ? "" : x.toFixed(2); }
  return String(v).trim().toLowerCase().replace(/\s+/g, " ");
}
// decide how to join rawA and rawB. approvedFields = array of field ids the supervisor OK'd (or null).
function keyPlan(rawA, rawB, approvedFields) {
  const aId = idxOf(rawA.headers, P.idany), bId = idxOf(rawB.headers, P.idany);
  if (aId != null && bId != null) return { mode: "id", a: [{ id: "id", idx: aId }], b: [{ id: "id", idx: bId }], labels: ["claim identifier"] };
  // composite: fields present in BOTH files
  let usable = COMPOSITE_FIELDS.map((f) => ({ f, ai: idxOf(rawA.headers, f.pats), bi: idxOf(rawB.headers, f.pats) })).filter((x) => x.ai != null && x.bi != null);
  if (approvedFields && approvedFields.length) { const set = new Set(approvedFields); const kept = usable.filter((x) => set.has(x.f.id)); if (kept.length === approvedFields.length) usable = kept; else return { mode: "composite", a: usable.map((x) => ({ id: x.f.id, idx: x.ai })), b: usable.map((x) => ({ id: x.f.id, idx: x.bi })), labels: usable.map((x) => x.f.label), approved: false, broken: true }; }
  if (!usable.length) return { mode: "none" };
  return { mode: "composite", a: usable.map((x) => ({ id: x.f.id, idx: x.ai })), b: usable.map((x) => ({ id: x.f.id, idx: x.bi })), labels: usable.map((x) => x.f.label), approved: !!(approvedFields && approvedFields.length) };
}
const rowKey = (row, cols) => cols.map((c) => normCell((row || [])[c.idx], c.id)).join("¦");
function runJoin(rawA, rawB, plan) {
  const mapA = new Map(); let dupA = 0;
  for (const r of rawA.rows) { const k = rowKey(r, plan.a); if (!k.replace(/¦/g, "")) continue; if (mapA.has(k)) dupA++; else mapA.set(k, r); }
  const pairs = []; let matched = 0;
  for (const r of rawB.rows) { const k = rowKey(r, plan.b); if (!k.replace(/¦/g, "")) continue; const a = mapA.get(k); if (a) { pairs.push({ a, b: r }); matched++; } }
  return { pairs, matched, totalA: rawA.rows.length, totalB: rawB.rows.length, dupA };
}

function autoCompute(reps, numDays, joinKeys) {
  const out = {}; const ar = reps.ar, den = reps.denial, pay = reps.payment, aud = reps.audit, esc = reps.escalation;
  // AR Aging = (>90-day balance) / total AR
  if (ar?.aging && ar?.totalAR) { let o = 0; for (const [b, v] of Object.entries(ar.aging)) if (is90plus(b)) o += v; out["ar.arAging"] = { value: o / ar.totalAR, src: ">90d balance ÷ total AR", inputs: { ">90d balance": Math.round(o), "total AR": Math.round(ar.totalAR) } }; }
  // Collection Rate (net) = payments / (charges - contractual adjustments)
  if (pay?.payments != null && pay?.charges != null && pay?.adjustments != null && (pay.charges - pay.adjustments) !== 0)
    out["ar.collectionRate"] = { value: pay.payments / (pay.charges - pay.adjustments), src: "payments ÷ (charges − adj)", inputs: { payments: Math.round(pay.payments), charges: Math.round(pay.charges), "contractual adj": Math.round(pay.adjustments) } };
  // Denial Rate = denial count / posting (payment) claim count
  if (den?.count != null && (pay?.claims || pay?.rows)) { const posted = pay?.claims ?? pay?.rows; out["ar.denialRate"] = { value: den.count / posted, src: "denied claims ÷ posted claims", inputs: { "denied claims": den.count, "posted claims": posted } }; }
  // Days in AR = total AR / (total payments / days in period), the team's agreed formula
  if (ar?.totalAR != null && pay?.payments && numDays) out["ar.daysInAR"] = { value: Math.round(ar.totalAR / (pay.payments / numDays)), src: "AR ÷ (payments ÷ days)", inputs: { "total AR": Math.round(ar.totalAR), "total payments": Math.round(pay.payments), "days in period": numDays, "average daily payment": Math.round(pay.payments / numDays) } };
  // Appeal % = appeals filed / denials received
  if (ar?.appeals != null && den?.count) out["ar.appeal"] = { value: ar.appeals / den.count, src: "appeals ÷ denials", inputs: { appeals: ar.appeals, denials: den.count } };
  // Quality = average Final Audit Score %
  if (aud?.score != null) out["ar.quality"] = { value: aud.score > 1 ? aud.score / 100 : aud.score, src: "avg Final Audit Score %", inputs: { "avg score": aud.score.toFixed(1) + "%" } };
  if (reps.subaudit?.score != null) out["submission.quality"] = { value: reps.subaudit.score > 1 ? reps.subaudit.score / 100 : reps.subaudit.score, src: "avg Final Audit Score %", inputs: { "avg score": reps.subaudit.score.toFixed(1) + "%" } };
  if (reps.ppaudit?.score != null) out["pp.quality"] = { value: reps.ppaudit.score > 1 ? reps.ppaudit.score / 100 : reps.ppaudit.score, src: "avg Final Audit Score %", inputs: { "avg score": reps.ppaudit.score.toFixed(1) + "%" } };
  // Escalations = count
  // Submission scope
  const sub = reps.submission, rej = reps.rejection, inf = reps.inflow;
  if (sub?.count) {
    if (rej?.count != null) out["submission.rejectionRate"] = { value: rej.count / sub.count, src: "rejected ÷ submitted", inputs: { rejected: rej.count, submitted: sub.count } };
    if (rej?.count != null && den?.count != null) out["submission.cleanClaim"] = { value: Math.max(0, (sub.count - rej.count - den.count)) / sub.count, src: "(submitted − rejected − denied) ÷ submitted", inputs: { submitted: sub.count, rejected: rej.count, denied: den.count } };
    if (inf?.count) out["submission.inflowVsSubmitted"] = { value: sub.count / inf.count, src: "submitted ÷ inflow", inputs: { submitted: sub.count, inflow: inf.count } };
  }
  // EV — Verification Completion % and Urgent vs Regular
  const ev = reps.ev;
  if (ev?.evTotal) out["ev.completionRate"] = { value: ev.evCompleted / ev.evTotal, src: "completed ÷ total verifications", inputs: { completed: ev.evCompleted, total: ev.evTotal } };
  if (ev?.evUrgent != null && ev?.evRegular != null) { const t = ev.evUrgent + ev.evRegular; out["ev.urgentVsRegular"] = { value: `${ev.evUrgent} urgent / ${ev.evRegular} regular` + (t ? ` (${Math.round(ev.evUrgent / t * 100)}% urgent)` : ""), text: true, src: "urgent vs regular counts", inputs: { urgent: ev.evUrgent, regular: ev.evRegular } }; }

  // Underpayment Rate = total (allowed − paid, where paid<allowed) ÷ total allowed
  const led = reps.ledger;
  if (led?.ledgerAllowed) out["pp.underpaymentRate"] = { value: led.ledgerUnder / led.ledgerAllowed, src: "underpaid $ ÷ total allowed", inputs: { "underpaid $": Math.round(led.ledgerUnder), "total allowed": Math.round(led.ledgerAllowed) } };
  // Variance Rate = discrepancy amount (could not post) ÷ total posting amount that came in
  const varf = reps.variance, pay2 = reps.payment;
  const recv = reps.postrecv;
  if (varf?.varianceAmt != null && recv?.postingIn) {
    const inputs = { "discrepancy (from file) $": Math.round(varf.varianceAmt), "received for posting $": Math.round(recv.postingIn) };
    if (pay2?.payments != null) inputs["cross-check, received minus posted $"] = Math.round(Math.max(0, recv.postingIn - pay2.payments));
    out["pp.varianceRate"] = { value: varf.varianceAmt / recv.postingIn, src: "discrepancy file ÷ total received for posting", inputs };
  }
  // Payment Adjustment Rate = non-contractual write-off (excl CO45) ÷ gross charges billed
  const nc = reps.noncontr;
  if (nc?.writeoff != null && pay2?.charges) out["pp.paymentAdjustmentRate"] = { value: nc.writeoff / pay2.charges, src: "write-off (excl CO45) ÷ gross charges", inputs: { "write-off $": Math.round(nc.writeoff), "gross charges": Math.round(pay2.charges) } };
  // Recoupment Rate = amount recouped ÷ total payment received
  const rc = reps.recoup;
  if (rc?.recoupAmt != null && pay2?.payments) out["pp.recoupmentRate"] = { value: rc.recoupAmt / pay2.payments, src: "amount recouped ÷ total payment received", inputs: { "recouped $": Math.round(rc.recoupAmt), "total payment received $": Math.round(pay2.payments) } };
  // Patient vs Insurance Collection ratio = patient collection ÷ insurance collection
  if (pay2?.patientPaid != null && pay2?.insurancePaid) { const ratio = pay2.patientPaid / pay2.insurancePaid; out["pp.patientVsInsurance"] = { value: `$${Math.round(pay2.patientPaid).toLocaleString()} patient : $${Math.round(pay2.insurancePaid).toLocaleString()} insurance  (ratio ${ratio.toFixed(2)})`, text: true, src: "patient collection ÷ insurance collection", inputs: { "patient $": Math.round(pay2.patientPaid), "insurance $": Math.round(pay2.insurancePaid), "ratio (patient ÷ insurance)": ratio.toFixed(2) } }; }

  // Payment Cycle = avg(paid date − submission date) across matched claims, via the match engine
  const ps = reps.paysub, pp2 = reps.paypaid;
  if (ps?.raw && pp2?.raw && ps.dateIdx != null && pp2.dateIdx != null) {
    const approved = joinKeys?.["pp.paymentCycle"]?.fields || null;
    const plan = keyPlan(ps.raw, pp2.raw, approved);
    if (plan.mode === "none") {
      out["pp.paymentCycle"] = { value: undefined, src: "no shared key to match files", inputs: {}, note: "no-key" };
    } else if (plan.mode === "composite" && !plan.approved) {
      out["pp.paymentCycle"] = { needsJoin: { pair: "pp.paymentCycle", labels: plan.labels, fields: plan.a.map((c) => c.id) } };
    } else {
      const j = runJoin(ps.raw, pp2.raw, plan);
      let sum = 0, cnt = 0, bad = 0;
      for (const pr of j.pairs) { const sd = parseDate((pr.a || [])[ps.dateIdx]), pd = parseDate((pr.b || [])[pp2.dateIdx]); if (!sd || !pd) { bad++; continue; } const days = Math.round((pd - sd) / 86400000); if (days < 0) { bad++; continue; } sum += days; cnt++; }
      if (cnt) out["pp.paymentCycle"] = { value: Math.round(sum / cnt), fmt: FMT.DAYS, src: "avg (paid − submission) days", inputs: { "matched claims": cnt, "match key": plan.mode === "id" ? "claim identifier" : plan.labels.join(" + "), "of B rows": `${j.matched}/${j.totalB}` + (bad ? `, ${bad} skipped` : "") } };
    }
  }

  // PCC — call handling (inbound picked/missed/dropped), callbacks, correspondence, fax/email
  const calls = reps.pcccalls, cb = reps.pcccallback, corr = reps.pcccorr, fax = reps.pccfax;
  if (calls?.callsTotal) {
    out["pcc.inflowVsPicked"] = { value: calls.callsPicked / calls.callsTotal, src: "picked ÷ total inflow calls", inputs: { "picked-up": calls.callsPicked, "total inflow": calls.callsTotal, missed: calls.callsMissed, dropped: calls.callsDropped } };
    out["pcc.missedCalls"] = { value: calls.callsMissed, src: "missed calls in log", inputs: { missed: calls.callsMissed, "total inflow": calls.callsTotal } };
    out["pcc.callsDropped"] = { value: calls.callsDropped, src: "dropped calls in log", inputs: { dropped: calls.callsDropped, "total inflow": calls.callsTotal } };
    out["pcc.dropRate"] = { value: calls.callsDropped / calls.callsTotal, src: "dropped ÷ total inflow calls", inputs: { dropped: calls.callsDropped, "total inflow": calls.callsTotal } };
  }
  if (cb?.cbTotal != null) {
    out["pcc.callbacksMade"] = { value: cb.cbDone, src: "completed callbacks in callback file", inputs: { "callbacks made": cb.cbDone, "callbacks logged": cb.cbTotal } };
    const missedDen = calls?.callsMissed ?? cb.cbTotal;
    if (missedDen) out["pcc.callbackRate"] = { value: cb.cbDone / missedDen, src: "callbacks made ÷ missed calls", inputs: { "callbacks made": cb.cbDone, "missed calls": missedDen, "missed source": calls?.callsMissed != null ? "call log" : "callback-file rows" } };
  }
  if (corr?.corrTotal) out["pcc.correspondence"] = { value: corr.corrDone / corr.corrTotal, src: "completed ÷ received", inputs: { completed: corr.corrDone, received: corr.corrTotal } };
  if (fax?.faxTotal) out["pcc.faxEmail"] = { value: fax.faxDone / fax.faxTotal, src: "completed ÷ received", inputs: { completed: fax.faxDone, received: fax.faxTotal } };

  // Escalations — split by department across scopes (one tracker per client)
  if (esc?.count != null) {
    for (const sk of Object.keys(DEPT_PATS)) {
      if (!SCOPE_BY_KEY[sk]?.criteria?.some((c) => c.key === "escalations")) continue;
      let cnt = null;
      if (esc.byDept) cnt = Object.entries(esc.byDept).filter(([d]) => DEPT_PATS[sk].test(d)).reduce((s, [, v]) => s + v, 0);
      else if (sk === "ar") cnt = esc.count;
      if (cnt != null) out[`${sk}.escalations`] = { value: cnt, src: esc.byDept ? "tracker · " + sk.toUpperCase() + " dept" : "tracker rows", inputs: { escalations: cnt } };
    }
  }
  return out;
}


/* ==================== bulk upload: file to slot matching ==================== */
// Drishti: bulk upload matcher.
// Decides which upload slot each dropped file belongs to, using three layers:
//   1. the file name,
//   2. the column headers inside the file,
//   3. a cross-check between the two.
// Nothing is ever placed on a guess. A file is only marked "certain" when the name
// points at exactly one slot AND the columns inside agree. Everything else is handed
// back for a human to confirm. These functions are mirrored verbatim inside App.jsx.

/* ---------- column patterns (same families the KPI engine already uses) ---------- */
const RX = {
  balance: [/ar\s*balance/i, /remaining\s*balance/i, /outstanding/i, /\bbalance\b/i, /amount\s*due/i, /open\s*amount/i],
  charge: [/total\s*charge/i, /charge\s*amount/i, /charges/i, /\bcharge\b/i, /billed\s*amount/i],
  bucket: [/aging\s*bucket/i, /\baging\b/i, /\bbucket\b/i, /age\s*category/i],
  dos: [/date\s*of\s*service/i, /\bdos\b/i, /service\s*date/i],
  score: [/final\s*audit\s*score/i, /audit\s*score/i, /final\s*score/i, /quality\s*score/i, /\bscore\b/i],
  claimid: [/claim\s*id/i, /claim\s*no/i, /claim\s*#/i, /claim\s*number/i, /\bclaim\b/i, /\bicn\b/i],
  payment: [/payment\s*amount/i, /\bpaid\s*amount/i, /amount\s*paid/i, /\bpayment\b/i, /\bpaid\b/i, /collected/i],
  adj: [/contractual\s*adjust/i, /contractual/i, /\badjustment\b/i, /\badj\b/i],
  allowed: [/allowed\s*amount/i, /\ballowed\b/i],
  received: [/amount\s*received\s*for\s*posting/i, /received\s*for\s*posting/i, /amount\s*received/i],
  discrepancy: [/discrepancy\s*amount/i, /\bdiscrepanc/i, /\bvariance\b/i, /unable\s*to\s*post/i],
  writeoff: [/write[\s-]*off/i, /\badjustment\s*amount/i, /non[\s-]*contractual/i],
  reason: [/reason\s*code/i, /\breason\b/i, /\bcarc\b/i],
  recouped: [/recoup/i, /taken\s*back/i, /\bretract/i],
  subdate: [/submission\s*date/i, /submitted\s*date/i, /date\s*submitted/i, /\bbill(ed)?\s*date/i],
  paiddate: [/paid\s*date/i, /payment\s*date/i, /date\s*paid/i, /check\s*date/i, /post\s*date/i],
  srno: [/\bsr\s*no/i, /serial/i, /ticket/i, /\bsr\b/i],
  dept: [/department/i, /\bdept\b/i],
  disposition: [/disposition/i, /call\s*status/i, /call\s*result/i, /\boutcome\b/i],
  urgency: [/urgent/i, /priority/i, /request\s*type/i, /verification\s*type/i],
  verifstatus: [/verification\s*status/i],
  completed: [/completed?/i, /\bstatus\b/i, /\bsent\b/i, /\bdone\b/i],
  denialcode: [/denial\s*code/i, /denied\s*amount/i, /\bdenial\b/i, /\bdenied\b/i],
  rejectreason: [/rejection\s*reason/i, /\brejection\b/i, /\brejected\b/i],
};
const has = (headers, pats) => (headers || []).some((h) => pats.some((p) => p.test(String(h || "").replace(/_+/g, " "))));
const hasAll = (headers, ...groups) => groups.every((g) => has(headers, g));

/* ---------- file-name normalising ---------- */
// Strip the extension, punctuation, dates and the usual file-name noise so that
// "ABC Clinic_Denial-Report_Sep 2026 (final v2).xlsx" reads as "denial report".
const NOISE = /\b(final|copy|new|old|updated|revised|draft|latest|version|v\d+|report|reports|file|files|sheet|data|xls|xlsx|csv|jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december|q[1-4]|fy\d*|20\d\d|\d{1,2})\b/g;
function normName(name) {
  return String(name || "")
    .replace(/\.[a-z0-9]+$/i, " ")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(NOISE, " ")
    .replace(/\s+/g, " ")
    .trim();
}
// The raw normalised name, noise words kept. Used where a stripped word still matters.
function rawName(name) {
  return String(name || "").replace(/\.[a-z0-9]+$/i, " ").toLowerCase().replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();
}

/* ---------- one rule per upload slot ----------
   strong : the name says this slot and no other. Carries a distinguishing word.
   weak   : the name is compatible with this slot but could fit a sibling too.
   veto   : the name contains a word that rules this slot out.
   content: reads the headers. 2 = unmistakable, 1 = compatible, 0 = contradicts.
   family : slots that look alike. Used to explain an ambiguity in plain words.
---------------------------------------------------------------------------- */
const SLOT_RULES = [
  { key: "ar", scope: "ar", family: "AR inventory",
    strong: [/\bar\b.*invent/, /invent.*\bar\b/, /receivab.*invent/, /\bar\s*aging/, /aging.*\bar\b/],
    weak: [/invent/, /aging/, /receivab/],
    veto: [/\bprev/, /\bprior/, /\blast\b/, /audit/, /denial/, /denied/, /escalat/],
    content: (h) => (has(h, RX.balance) && (has(h, RX.bucket) || has(h, RX.dos)) ? 2 : has(h, RX.balance) ? 1 : 0) },

  { key: "arprev", scope: "ar", family: "AR inventory",
    strong: [/invent.*(prev|prior|last)/, /(prev|prior|last).*invent/, /\bar\b.*(prev|prior|last)/, /(prev|prior|last).*\bar\b/],
    weak: [], veto: [/audit/],
    content: (h) => (has(h, RX.balance) && (has(h, RX.bucket) || has(h, RX.dos)) ? 2 : has(h, RX.balance) ? 1 : 0) },

  { key: "denial", scope: "ar", family: null,
    strong: [/denial/, /denied/], weak: [], veto: [/audit/, /appeal/],
    content: (h) => (has(h, RX.denialcode) ? 2 : has(h, RX.claimid) ? 1 : 0) },

  { key: "audit", scope: "ar", family: "audit",
    strong: [/\bar\b.*audit/, /audit.*\bar\b/, /receivab.*audit/],
    weak: [], veto: [/submiss/, /posting/, /payment/],
    content: (h) => (has(h, RX.score) ? 2 : 0) },

  { key: "escalation", scope: "ar", family: null,
    strong: [/escalat/], weak: [], veto: [],
    content: (h) => (has(h, RX.srno) || has(h, RX.dept) ? 2 : 1) },

  { key: "submission", scope: "submission", family: "submission",
    strong: [/claim\s*submission/, /submission\s*claim/],
    weak: [/submiss/, /submitted/, /\bbilled\b/],
    veto: [/audit/, /inflow/, /cycle/, /reject/, /denial/],
    content: (h) => (has(h, RX.claimid) && has(h, RX.subdate) ? 2 : has(h, RX.claimid) ? 1 : 0) },

  { key: "rejection", scope: "submission", family: null,
    strong: [/reject/], weak: [], veto: [/audit/],
    content: (h) => (has(h, RX.rejectreason) ? 2 : has(h, RX.claimid) ? 1 : 0) },

  { key: "inflow", scope: "submission", family: null,
    strong: [/inflow/], weak: [], veto: [/audit/, /\bcall\b/, /\bpcc\b/, /picked/],
    content: (h) => (has(h, RX.claimid) ? 1 : 0) },

  { key: "subaudit", scope: "submission", family: "audit",
    strong: [/submiss.*audit/, /audit.*submiss/, /claim.*audit/],
    weak: [], veto: [/posting/, /\bar\b/, /receivab/],
    content: (h) => (has(h, RX.score) ? 2 : 0) },

  { key: "payment", scope: "pp", family: "payment posting",
    strong: [/payment\s*posting/, /posting\s*payment/, /\bera\b/],
    weak: [/posting/, /payment/],
    veto: [/audit/, /cycle/, /ledger/, /recoup/, /varian/, /discrep/, /non\s*contract/, /write\s*off/, /received\s*for/],
    content: (h) => (hasAll(h, RX.charge, RX.adj, RX.payment) ? 2 : has(h, RX.payment) ? 1 : 0) },

  { key: "paysub", scope: "pp", family: "payment cycle",
    strong: [/submiss.*cycle/, /cycle.*submiss/, /submiss.*payment\s*cycle/],
    weak: [], veto: [/paid/],
    content: (h) => (has(h, RX.claimid) && has(h, RX.subdate) ? 2 : has(h, RX.claimid) ? 1 : 0) },

  { key: "paypaid", scope: "pp", family: "payment cycle",
    strong: [/paid\s*claim/, /claim.*paid/, /paid.*cycle/, /cycle.*paid/],
    weak: [], veto: [/submiss/],
    content: (h) => (has(h, RX.claimid) && has(h, RX.paiddate) ? 2 : has(h, RX.claimid) ? 1 : 0) },

  { key: "ledger", scope: "pp", family: null,
    strong: [/ledger/], weak: [], veto: [],
    content: (h) => (hasAll(h, RX.allowed, RX.payment) ? 2 : has(h, RX.allowed) ? 1 : 0) },

  { key: "postrecv", scope: "pp", family: null,
    strong: [/received\s*for\s*posting/, /amount\s*received/, /posting\s*inflow/],
    weak: [], veto: [/varian/, /discrep/],
    content: (h) => (has(h, RX.received) ? 2 : 0) },

  { key: "variance", scope: "pp", family: null,
    strong: [/varian/, /discrep/], weak: [], veto: [],
    content: (h) => (has(h, RX.discrepancy) ? 2 : 0) },

  { key: "noncontr", scope: "pp", family: null,
    strong: [/non\s*contract/, /write\s*off/, /\btfl\b/], weak: [], veto: [],
    content: (h) => (has(h, RX.writeoff) && has(h, RX.reason) ? 2 : has(h, RX.writeoff) ? 1 : 0) },

  { key: "recoup", scope: "pp", family: null,
    strong: [/recoup/, /taken\s*back/, /retract/], weak: [], veto: [],
    content: (h) => (has(h, RX.recouped) ? 2 : 0) },

  { key: "ppaudit", scope: "pp", family: "audit",
    strong: [/posting.*audit/, /audit.*posting/, /payment.*audit/, /audit.*payment/],
    weak: [], veto: [/\bar\b/, /receivab/, /submiss/],
    content: (h) => (has(h, RX.score) ? 2 : 0) },

  { key: "ev", scope: "ev", family: null,
    strong: [/benefit\s*verif/, /verif.*benefit/, /\bev\b/, /eligib/],
    weak: [/verif/], veto: [],
    content: (h) => (has(h, RX.verifstatus) || (has(h, RX.urgency) && has(h, RX.completed)) ? 2 : has(h, RX.completed) ? 1 : 0) },

  { key: "pcccalls", scope: "pcc", family: null,
    strong: [/call\s*log/, /inbound\s*call/, /call.*disposition/],
    weak: [/\bcalls?\b/], veto: [/callback/, /call\s*back/, /drop/],
    content: (h) => (has(h, RX.disposition) ? 2 : 0) },

  { key: "pcccallback", scope: "pcc", family: null,
    strong: [/callback/, /call\s*back/], weak: [], veto: [],
    content: (h) => (has(h, RX.completed) ? 1 : 0) },

  { key: "pcccorr", scope: "pcc", family: null,
    strong: [/correspond/], weak: [], veto: [],
    content: (h) => (has(h, RX.completed) ? 1 : 0) },

  { key: "pccfax", scope: "pcc", family: null,
    strong: [/\bfax\b/, /\bemail\b/, /\bmail\b/], weak: [], veto: [/correspond/],
    content: (h) => (has(h, RX.completed) ? 1 : 0) },
];
const RULE_BY_KEY = Object.fromEntries(SLOT_RULES.map((r) => [r.key, r]));

/* ---------- layer 1: what does the name say ---------- */
function nameScore(fileName, rule) {
  const s = normName(fileName), raw = rawName(fileName);
  const hit = (pats) => pats.some((p) => p.test(s) || p.test(raw));
  if (rule.veto.length && hit(rule.veto)) return 0;
  if (rule.strong.length && hit(rule.strong)) return 2;
  if (rule.weak.length && hit(rule.weak)) return 1;
  return 0;
}

/* ---------- layer 2 + 3: decide, and refuse to guess ---------- */
// Returns one row per file:
//   status "certain"  name points at one slot and the columns agree
//   status "check"    a single sensible candidate, but something is short of proof
//   status "unknown"  nothing reliable to go on
//   slot              the proposed slot, or null
//   options           slots worth offering in the dropdown
//   reason            plain words for the review table
function matchOne(file, allowedKeys) {
  const headers = file.headers || [];
  const rules = SLOT_RULES.filter((r) => !allowedKeys || allowedKeys.includes(r.key));
  const scored = rules.map((r) => ({ key: r.key, rule: r, name: nameScore(file.name, r), content: r.content(headers) }));

  const byName = scored.filter((x) => x.name > 0);
  const strong = byName.filter((x) => x.name === 2);
  const weak = byName.filter((x) => x.name === 1);
  const contentSure = scored.filter((x) => x.content === 2);

  const opts = (list) => list.map((x) => x.key);

  // the name names exactly one slot
  if (strong.length === 1) {
    const c = strong[0];
    const rival = contentSure.filter((x) => x.key !== c.key);
    if (c.content === 2) return { slot: c.key, status: "certain", reason: "Name and columns both match", options: opts(scored.filter((x) => x.content > 0 || x.name > 0)) };
    // a merely compatible column set is not proof while another slot's columns are unmistakable
    if (c.content === 1 && !rival.length) return { slot: c.key, status: "certain", reason: "Name matches, columns are consistent", options: opts(scored.filter((x) => x.content > 0 || x.name > 0)) };
    if (c.content === 1) return { slot: c.key, status: "check", reason: `Name says this, but the columns look like the ${RULE_BY_KEY[rival[0].key] ? rival[0].key : "another"} file`, options: opts([c, ...rival]) };
    // the name is clear but the columns are not what this slot needs: do not place it
    return { slot: c.key, status: "check", reason: rival.length ? `Name says this, but the columns look like the ${rival[0].key} file` : "Name matches, but the expected columns were not found", options: opts([c, ...rival]) };
  }

  // more than one slot claims the name
  if (strong.length > 1) {
    const agree = strong.filter((x) => x.content >= 1);
    if (agree.length === 1) return { slot: agree[0].key, status: "check", reason: `Name could be ${strong.map((x) => x.key).join(" or ")}; columns point to this one`, options: opts(strong) };
    const fam = strong[0].rule.family;
    return { slot: null, status: "unknown", reason: fam ? `Name could be any of the ${fam} files` : "Name matches more than one slot", options: opts(strong) };
  }

  // only loose name matches
  if (weak.length === 1) {
    const c = weak[0];
    if (c.content === 2) return { slot: c.key, status: "certain", reason: "Columns are unmistakable and the name agrees", options: opts(scored.filter((x) => x.content > 0 || x.name > 0)) };
    if (c.content === 1) return { slot: c.key, status: "check", reason: "Name is close, columns are consistent", options: opts(scored.filter((x) => x.content > 0 || x.name > 0)) };
    return { slot: null, status: "unknown", reason: "Name is a loose match and the columns do not confirm it", options: opts([c, ...contentSure]) };
  }
  if (weak.length > 1) {
    const agree = weak.filter((x) => x.content === 2);
    if (agree.length === 1) return { slot: agree[0].key, status: "check", reason: "Name is vague; columns point to this one", options: opts(weak) };
    return { slot: null, status: "unknown", reason: "Name is a loose match for more than one slot", options: opts(weak) };
  }

  // nothing in the name: fall back to the columns alone, and never call it certain
  if (contentSure.length === 1) return { slot: contentSure[0].key, status: "check", reason: "Name gave nothing away; the columns suggest this", options: opts(contentSure) };
  if (contentSure.length > 1) {
    const fams = [...new Set(contentSure.map((x) => x.rule.family).filter(Boolean))];
    const why = fams.length === 1 && contentSure.every((x) => x.rule.family) ? `Could be any of the ${fams[0]} files, so the name needs to say which` : "Columns fit more than one slot";
    return { slot: null, status: "unknown", reason: why, options: opts(contentSure) };
  }
  return { slot: null, status: "unknown", reason: "Not recognised", options: [] };
}

// Match a whole drop. Two files landing on one slot is itself a mismatch risk,
// so both are demoted and handed back rather than one silently overwriting the other.
function matchFiles(files, allowedKeys) {
  const rows = (files || []).map((f, i) => {
    const r = { i, name: f.name, headers: f.headers || [], ...matchOne(f, allowedKeys) };
    // recognised, but it belongs to a scope this client does not have: say so instead of "not recognised"
    if (allowedKeys && r.status !== "certain") {
      const wide = matchOne(f, null);
      if (wide.status === "certain" && !allowedKeys.includes(wide.slot)) { r.slot = null; r.status = "unknown"; r.outOfScope = wide.slot; r.reason = "Belongs to a scope this client does not have"; }
    }
    return r;
  });
  const bySlot = {};
  for (const r of rows) if (r.slot) (bySlot[r.slot] = bySlot[r.slot] || []).push(r);
  for (const [slot, list] of Object.entries(bySlot)) {
    if (list.length < 2) continue;
    for (const r of list) { r.status = "conflict"; r.reason = `${list.length} files claim this slot, pick one`; if (!r.options.includes(slot)) r.options = [slot, ...r.options]; }
  }
  return rows;
}

const summarise = (rows) => ({
  certain: rows.filter((r) => r.status === "certain").length,
  check: rows.filter((r) => r.status === "check").length,
  conflict: rows.filter((r) => r.status === "conflict").length,
  unknown: rows.filter((r) => r.status === "unknown").length,
  total: rows.length,
});
// Only rows with a slot chosen are applied, and a conflict is never applied until resolved.
const readyRows = (rows) => rows.filter((r) => r.slot && r.status !== "conflict" && !r.skip);

/* ============================== styles ================================= */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');
:root{--ink:#16233b;--ink-2:#4a5a72;--ink-3:#8a97ab;--line:#e4e8ee;--line-2:#eef1f5;--bg:#eef2f6;--surface:#fff;--navy:#16233b;--navy-2:#22385a;--steel:#0f9f88;--steel-soft:#e7f6f3;--brand1:#3ECF7C;--brand2:#2CC9B0;--brand3:#38C8E8;--great:#0b7a4f;--great-bg:#d3efe1;--ok:#1f9d6b;--ok-bg:#e7f5ee;--warn:#c07d10;--warn-bg:#fbf1dc;--miss:#cf3d3d;--miss-bg:#fbe8e8;--info:#64748b;--info-bg:#eef1f4;--grad:linear-gradient(120deg,var(--brand1),var(--brand2) 52%,var(--brand3));--sh-1:0 1px 2px rgba(22,35,59,.05);--sh-2:0 1px 3px rgba(22,35,59,.05),0 10px 26px -12px rgba(22,35,59,.16);--sh-3:0 8px 20px -6px rgba(22,35,59,.14),0 24px 54px -18px rgba(20,120,110,.18);--ease:cubic-bezier(.4,.14,.2,1)}
*{box-sizing:border-box}
@keyframes riseIn{0%{opacity:0;transform:translateY(8px)}100%{opacity:1;transform:none}}
@keyframes shimmer{0%{background-position:0% 50%}100%{background-position:200% 50%}}
.app{font-family:'Poppins',system-ui,sans-serif;color:var(--ink);background:radial-gradient(120% 80% at 8% -8%,#e4f7ef 0%,rgba(228,247,239,0) 46%),radial-gradient(120% 80% at 98% 0%,#e2f4fb 0%,rgba(226,244,251,0) 42%),linear-gradient(180deg,#f4fbf9 0%,#eef4f5 55%,#eaf0f3 100%);background-attachment:fixed;min-height:100vh;font-size:14px;-webkit-font-smoothing:antialiased}
.mono{font-family:'Poppins',ui-monospace,monospace;font-feature-settings:"tnum" 1}
.auth-wrap{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;background:radial-gradient(130% 90% at 50% -10%,#e0f7ef 0%,#eef3f7 42%,#e7ecf1 100%)}
.auth{background:var(--surface);border:1px solid var(--line);border-radius:18px;width:100%;max-width:400px;padding:32px 32px 26px;box-shadow:var(--sh-3);position:relative;overflow:hidden;animation:riseIn .5s var(--ease) both}
.auth::before{content:"";position:absolute;top:0;left:0;right:0;height:4px;background:var(--grad)}
.auth .brand,.nav .brand{display:flex;align-items:center;gap:10px;font-weight:600;letter-spacing:-.01em}
.auth .brand{font-size:17px;margin-bottom:4px}
.auth .brand .logo,.nav .brand .logo{display:block;flex:none;filter:drop-shadow(0 3px 6px rgba(44,201,176,.28))}
.brand{display:flex;align-items:center;gap:9px}
.brandtext{display:flex;flex-direction:column;line-height:1.02}
.bytag{font-size:9.5px;font-weight:500;color:var(--ink-3);letter-spacing:.02em}
.nav .brand{font-size:15px}
.wm{font-weight:700;letter-spacing:-.02em}
.wm-navy{color:var(--navy)}
.wm-grad{background:linear-gradient(90deg,var(--brand1),var(--brand2),var(--brand3));background-size:200% auto;-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;color:transparent}
.auth .sub{color:var(--ink-3);font-size:12.5px;margin-bottom:22px}
.auth h2{font-size:15px;margin:0 0 16px}
.nav{background:rgba(255,255,255,.78);backdrop-filter:saturate(1.5) blur(14px);-webkit-backdrop-filter:saturate(1.5) blur(14px);border-bottom:1px solid rgba(228,232,238,.8);padding:0 24px;display:flex;align-items:center;gap:5px;height:62px;position:sticky;top:0;z-index:20;flex-wrap:wrap;box-shadow:0 1px 0 rgba(22,35,59,.02),0 8px 26px -18px rgba(22,35,59,.28)}
.nav{position:sticky}
.nav::after{content:"";position:absolute;left:0;right:0;bottom:-1px;height:3px;background-image:linear-gradient(90deg,transparent,rgba(255,255,255,.95) 50%,transparent),linear-gradient(90deg,transparent,var(--brand1) 18%,var(--brand2) 50%,var(--brand3) 82%,transparent);background-size:28% 100%,100% 100%;background-repeat:no-repeat;background-position:-60% 0,0 0;animation:streak 3.2s linear infinite;pointer-events:none}
.nav .brand{margin-right:20px}
.tab{border:none;background:none;font-family:inherit;font-size:13.5px;color:var(--ink-2);padding:9px 16px;border-radius:999px;cursor:pointer;position:relative;transition:background .18s var(--ease),color .18s var(--ease),box-shadow .18s var(--ease)}
.tab:hover{background:var(--line-2);color:var(--ink)}.tab.on{background:linear-gradient(135deg,rgba(62,207,124,.16),rgba(56,200,232,.16));color:var(--steel);font-weight:600;box-shadow:inset 0 0 0 1px rgba(44,201,176,.32)}
.nav .spacer{flex:1}.who{display:flex;align-items:center;gap:10px;font-size:13px;color:var(--ink-2)}
.rolechip{font-size:10.5px;font-weight:600;padding:3px 9px;border-radius:20px;background:var(--steel-soft);color:var(--steel);text-transform:capitalize}
.wrap{max-width:1180px;margin:0 auto;padding:24px 24px 70px;animation:riseIn .4s var(--ease) both}
.pagehead{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-bottom:18px}
.pagehead h1{font-size:25px;font-weight:700;letter-spacing:-.03em;margin:0}.pagehead .sub{color:var(--ink-3);font-size:12.5px;margin-top:5px}
.row{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
select,input,textarea{font-family:inherit;font-size:13.5px;padding:9px 13px;border:1px solid var(--line);border-radius:12px;background:var(--surface);color:var(--ink);transition:border-color .16s var(--ease),box-shadow .16s var(--ease)}
select:focus,input:focus,textarea:focus{outline:none;border-color:var(--steel);box-shadow:0 0 0 3px rgba(15,159,136,.14)}
.btn{font-family:inherit;font-size:13px;font-weight:500;padding:9px 17px;border-radius:999px;border:1px solid var(--line);background:var(--surface);color:var(--ink);cursor:pointer;transition:transform .14s var(--ease),box-shadow .16s var(--ease),border-color .16s var(--ease),filter .16s var(--ease)}
.btn:hover{border-color:var(--ink-3);box-shadow:var(--sh-1);transform:translateY(-1px)}.btn:active{transform:translateY(0)}
.btn.primary{background:var(--grad);background-size:160% auto;border:none;color:#fff;font-weight:600;box-shadow:0 6px 16px -6px rgba(44,201,176,.6)}.btn.primary:hover{filter:brightness(1.04);background-position:80% 50%;box-shadow:0 10px 22px -6px rgba(44,201,176,.7)}
.btn.primary:disabled{opacity:.5;cursor:not-allowed;box-shadow:none;transform:none}.btn.sm{padding:6px 11px;font-size:12px}.btn.ghost{border-color:transparent;color:var(--ink-3);box-shadow:none}.btn.ghost:hover{color:var(--miss);background:var(--miss-bg);border-color:transparent}
.btn.danger{color:var(--miss);border-color:rgba(207,61,61,.35)}
.btn.danger:hover{background:var(--miss-bg);border-color:var(--miss);box-shadow:var(--sh-1)}
.btn.danger.solid{background:linear-gradient(135deg,#e05858,#cf3d3d);color:#fff;border:none;font-weight:600;box-shadow:0 6px 16px -6px rgba(207,61,61,.6)}
.btn.danger.solid:hover{filter:brightness(1.05)}
.btn.danger.solid:disabled{opacity:.45;cursor:not-allowed;box-shadow:none;transform:none;filter:none}
.link{background:none;border:none;color:var(--steel);font-family:inherit;font-size:12.5px;cursor:pointer;padding:0;text-decoration:underline}
.strip{background:linear-gradient(135deg,#fff 0%,#f1fbf7 55%,#eaf8fb 100%);border:1px solid var(--line);border-radius:24px;padding:22px 26px;display:flex;align-items:center;gap:24px;flex-wrap:wrap;margin-bottom:12px;box-shadow:var(--sh-2);position:relative;overflow:hidden}
.strip::before{content:"";position:absolute;top:0;left:0;bottom:0;width:5px;background:var(--grad)}
.strip .headline{font-size:15px;font-weight:600}.strip .headline b{font-family:'Poppins'}
.bar{flex:1;min-width:170px;height:10px;border-radius:6px;overflow:hidden;display:flex;background:var(--line-2);box-shadow:inset 0 1px 2px rgba(22,35,59,.08)}.bar>span{height:100%;transition:width .6s var(--ease)}
.tallies{display:flex;gap:10px;flex-wrap:wrap}
.tally{display:flex;align-items:center;gap:7px;padding:8px 13px;border-radius:10px;font-size:13px;font-weight:500;transition:transform .14s var(--ease)}
.tally:hover{transform:translateY(-1px)}
.tally .n{font-family:'Poppins';font-weight:700;font-size:15px}
.tally.ok{background:var(--ok-bg);color:var(--ok)}.tally.warn{background:var(--warn-bg);color:var(--warn)}.tally.miss{background:var(--miss-bg);color:var(--miss)}
.section-h{margin:28px 0 14px;display:flex;align-items:center;gap:12px}
.section-h h2{font-size:13px;font-weight:600;color:var(--navy-2);margin:0;letter-spacing:-.01em;display:flex;align-items:center;gap:9px;background:rgba(255,255,255,.7);padding:7px 15px 7px 11px;border-radius:999px;box-shadow:0 1px 2px rgba(22,35,59,.05),inset 0 0 0 1px rgba(228,232,238,.7)}
.section-h h2::before{content:"";width:12px;height:12px;border-radius:50%;background:var(--grad);box-shadow:0 2px 6px -1px rgba(44,201,176,.55)}
.section-h .rule{display:none}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:14px}
.tile{background:var(--surface);border:1px solid var(--line);border-radius:20px;padding:17px 18px;border-left-width:4px;box-shadow:var(--sh-1);transition:transform .16s var(--ease),box-shadow .18s var(--ease);animation:riseIn .4s var(--ease) both}
.tile:hover{transform:translateY(-3px);box-shadow:var(--sh-3)}
.tile.great{border-left-color:var(--great)}.tile.ok{border-left-color:var(--ok)}.tile.warn{border-left-color:var(--warn)}.tile.miss{border-left-color:var(--miss)}.tile.info{border-left-color:var(--info)}
.tile .m-top{display:flex;justify-content:space-between;align-items:flex-start;gap:8px}
.tile .m-name{font-size:13px;font-weight:600;color:var(--ink-2)}
.chip{font-size:10.5px;font-weight:600;padding:3px 9px;border-radius:20px;white-space:nowrap}
.chip.great{background:var(--great-bg);color:var(--great)}.chip.ok{background:var(--ok-bg);color:var(--ok)}.chip.warn{background:var(--warn-bg);color:var(--warn)}.chip.miss{background:var(--miss-bg);color:var(--miss)}.chip.info{background:var(--info-bg);color:var(--info)}
.tile .m-val{font-family:'Poppins';font-size:28px;font-weight:700;letter-spacing:-.02em;margin:10px 0 2px;line-height:1.05}
.tile .m-val.txt{font-size:16px;font-family:'Poppins';font-weight:600}
.tile .m-target{font-size:11.5px;color:var(--ink-3)}
.tile .m-foot{display:flex;justify-content:space-between;align-items:center;margin-top:11px;padding-top:10px;border-top:1px solid var(--line-2)}
.tile .m-note{font-size:11px;color:var(--ink-3)}
.tile .delta{font-family:'Poppins';font-size:11.5px;font-weight:500}
.delta.up{color:var(--ok)}.delta.down{color:var(--miss)}.delta.flat{color:var(--ink-3)}
.gear{border:none;background:none;color:var(--ink-3);cursor:pointer;font-size:13px;padding:2px 4px;border-radius:5px;transition:background .16s,color .16s}.gear:hover{background:var(--line-2);color:var(--steel)}
.scope-block{background:var(--surface);border:1px solid var(--line);border-radius:20px;margin-bottom:14px;overflow:hidden;box-shadow:var(--sh-1)}
.scope-head{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 18px;cursor:pointer;background:linear-gradient(135deg,#f7fbfa,var(--line-2));transition:background .18s var(--ease)}
.scope-head:hover{background:linear-gradient(135deg,#eef8f5,#e9f5f1)}
.scope-head .name{font-weight:600;font-size:14px}
.scope-head .meta{font-size:12px;color:var(--ink-3)}
.scope-body{padding:6px 18px 16px}
.crow{display:grid;grid-template-columns:1.4fr 1.3fr 1fr auto;gap:14px;align-items:center;padding:11px 0;border-bottom:1px solid var(--line-2)}
.crow:last-child{border-bottom:none}
.crow .cn{font-weight:600;font-size:13px}.crow .cstd{font-size:11.5px;color:var(--ink-3);margin-top:2px}
.crow .cin input,.crow .cin select{width:100%}
.crow .cin.pct{position:relative}
.crow .cres{display:flex;align-items:center;gap:8px;justify-content:flex-end}
.crow .cres .v{font-family:'Poppins';font-weight:600;font-size:14px}
.calc-inline{grid-column:1 / -1;background:var(--steel-soft);border-radius:9px;padding:12px 14px;margin-top:-2px}
.calc-inline .fields{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px}
.calc-inline label{display:block;font-size:11.5px;color:var(--ink-2);margin-bottom:4px;font-weight:500}
.calc-inline input{width:100%;font-family:'Poppins'}
.avg3{display:flex;gap:6px;margin-top:5px}.avg3 input{width:100%;padding:6px 8px;font-size:12px}
.hintline{font-size:11px;color:var(--ink-3);margin-top:5px}
.upl{display:flex;align-items:center;gap:10px;margin:6px 0 2px;flex-wrap:wrap}
.mapper{background:#fff;border:1px solid #d6e3ea;border-radius:9px;padding:11px 13px;margin-top:8px}
.mapper .mh{font-size:12px;font-weight:600;color:var(--steel);margin-bottom:8px}
.maprow{display:grid;grid-template-columns:1fr 1.2fr;gap:10px;align-items:center;margin-bottom:7px}.maprow .lab{font-size:12px;color:var(--ink-2)}
.panel{background:var(--surface);border:1px solid var(--line);border-radius:18px;overflow:hidden;box-shadow:var(--sh-1)}
tbody tr{transition:background .14s var(--ease)}tbody tr:hover td{background:rgba(15,159,136,.035)}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;font-size:11px;color:var(--ink-3);font-weight:600;padding:10px 14px;border-bottom:1px solid var(--line);background:var(--line-2)}
td{padding:9px 14px;border-bottom:1px solid var(--line-2);color:var(--ink-2)}tr:last-child td{border-bottom:none}
td input{width:100%;padding:6px 9px}
.tasktbl td .mono{font-family:'Poppins'}
.tag{font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:20px;background:var(--info-bg);color:var(--info)}.tag.up{background:var(--steel-soft);color:var(--steel)}.tag.man{background:var(--warn-bg);color:var(--warn)}
.krow{display:grid;grid-template-columns:1.7fr 128px 66px 88px;gap:12px;align-items:center;padding:11px 0;border-bottom:1px solid var(--line-2)}
.krow:not(.khead){padding-left:11px;padding-right:11px;margin-left:-11px;margin-right:-11px;border-radius:12px;transition:background .15s var(--ease)}
.krow:not(.khead):hover{background:linear-gradient(90deg,rgba(15,159,136,.07),rgba(56,200,232,.04))}
.krow.khead{font-size:10.5px;color:var(--ink-3);text-transform:uppercase;letter-spacing:.03em;padding:4px 0}
.krow .cn{font-size:13px;font-weight:600}
.infobtn{border:none;background:none;cursor:pointer;color:var(--ink-3);font-size:12px;margin-left:6px;padding:0;vertical-align:middle}
.infobtn:hover{color:var(--steel)}
.krow .cstd{font-size:11px;color:var(--ink-3);margin-top:2px}
.krow .cin input,.krow .cin select{width:100%}
.unitbox{display:flex;align-items:center;gap:4px}
.unitbox input{flex:1;min-width:0;text-align:right}
.unitbox i{font-style:normal;font-size:11px;color:var(--ink-3)}
.krow .csrc{font-size:11px}
.reasonrow{padding:0 0 10px 0;margin-top:-4px}
.reasonrow input{width:100%;font-size:12px;border-color:var(--warn);background:var(--warn-bg)}
.seg{display:inline-flex;background:var(--line-2);border-radius:11px;padding:4px;gap:3px}
.segbtn{border:none;background:none;font-family:inherit;font-size:13px;font-weight:500;padding:7px 17px;border-radius:8px;cursor:pointer;color:var(--ink-2);transition:background .16s var(--ease),color .16s var(--ease)}
.segbtn:hover{color:var(--ink)}
.segbtn.on{background:var(--surface);color:var(--steel);box-shadow:0 1px 3px rgba(24,28,34,.12);font-weight:600}
.flab{display:block;font-size:11px;color:var(--ink-3);margin-bottom:4px}
.group-h{font-size:15px;font-weight:600;color:var(--navy);margin:22px 0 4px;padding-bottom:6px;border-bottom:2px solid var(--brand2);letter-spacing:.02em}
.group-h:first-child{margin-top:0}
.cmptbl{width:100%;border-collapse:collapse;font-size:12.5px}
.cmptbl th{text-align:left;padding:8px 14px;font-weight:500;color:var(--ink-3);font-size:10.5px;text-transform:uppercase;letter-spacing:.03em;background:var(--line-2)}
.cmptbl td{padding:7px 14px;border-top:1px solid var(--line-2)}
.cmptbl tr.tot td{font-weight:600;border-top:2px solid var(--line)}
.cmptbl tr.unassigned-row td{background:var(--warn-bg)}
.cmptbl tr.cnt td{font-size:11px;color:var(--ink-3);border-top:1px solid var(--line-2)}
.cmp.ok{color:var(--ok);font-weight:600}
.cmp.miss{color:var(--miss);font-weight:600}
.cmptbl.matx{min-width:820px}
.cmptbl.matx th,.cmptbl.matx td{padding:6px 10px;font-size:11.5px}
.empty{display:flex;flex-direction:column;align-items:center;text-align:center;padding:64px 20px;color:var(--ink-3)}
.card{background:var(--surface);border:1px solid var(--line);border-radius:20px;padding:20px 22px;box-shadow:var(--sh-1)}
.scopepick{display:flex;flex-direction:column;gap:8px;margin:4px 0 2px}
.scopepick label{display:flex;align-items:flex-start;gap:10px;padding:11px 13px;border:1px solid var(--line);border-radius:11px;cursor:pointer;font-size:13px;transition:border-color .16s var(--ease),background .16s var(--ease),transform .12s var(--ease)}
.scopepick label:hover{border-color:var(--brand2);transform:translateY(-1px)}
.scopepick label.on{border-color:var(--steel);background:var(--steel-soft);box-shadow:0 0 0 3px rgba(15,159,136,.12)}
.scopepick input{margin-top:2px}
.scopepick .d{font-size:11.5px;color:var(--ink-3);margin-top:2px}
.overlay{font-family:'Poppins',system-ui,sans-serif;color:var(--ink);position:fixed;inset:0;background:rgba(16,26,42,.5);backdrop-filter:blur(3px);-webkit-backdrop-filter:blur(3px);display:flex;align-items:center;justify-content:center;padding:20px;z-index:50;animation:riseIn .22s var(--ease) both}
.modal{background:var(--surface);border-radius:18px;width:100%;max-width:460px;padding:24px 26px;box-shadow:0 30px 80px -20px rgba(16,26,42,.5);max-height:88vh;overflow-y:auto;position:relative;animation:riseIn .3s var(--ease) both}
.modal::before{content:"";position:absolute;top:0;left:0;right:0;height:4px;border-radius:18px 18px 0 0;background:var(--grad)}
.modal h3{margin:0 0 4px;font-size:16px}.modal p{margin:0 0 16px;color:var(--ink-3);font-size:13px}
.field{margin-bottom:13px}.field label{display:block;font-size:12px;color:var(--ink-2);margin-bottom:5px;font-weight:500}
.field input,.field select{width:100%}.two{display:grid;grid-template-columns:1fr 1fr;gap:11px}
.modal-actions{display:flex;justify-content:space-between;gap:10px;margin-top:20px}
.auth.doors-wide{max-width:580px}.auth.wide{max-width:490px}
.doors{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.door{text-align:left;display:flex;flex-direction:column;align-items:flex-start;gap:7px;padding:18px 16px;border:1px solid var(--line);border-radius:14px;background:var(--surface);cursor:pointer;font-family:inherit;color:var(--ink);transition:border-color .16s var(--ease),box-shadow .16s var(--ease),transform .16s var(--ease)}
.door:hover{border-color:var(--brand2);box-shadow:var(--sh-2);transform:translateY(-1px)}
.door-ic{width:34px;height:34px;border-radius:10px;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,var(--brand1),var(--brand3));color:#fff;font-size:14px;box-shadow:0 3px 8px -2px rgba(44,201,176,.5)}
.door-t{font-weight:600;font-size:14.5px}.door-d{font-size:12px;color:var(--ink-3);line-height:1.5}
@media (max-width:600px){.doors{grid-template-columns:1fr}}
.auth-links{display:flex;justify-content:space-between;gap:10px;margin-top:14px}
.auth-lead{color:var(--ink-3);font-size:12.5px;margin:-8px 0 16px}
.auth-seg{display:flex;width:100%;margin-bottom:18px}.auth-seg .segbtn{flex:1}
.req{font-size:9.5px;font-weight:600;color:var(--miss);text-transform:uppercase;letter-spacing:.05em;margin-left:5px}
.probbar{display:flex;align-items:center;justify-content:space-between;gap:12px;margin:12px auto 0;max-width:1240px;width:calc(100% - 32px);background:var(--miss-bg);color:var(--miss);border:1px solid #f1c9c9;border-radius:10px;padding:9px 14px;font-size:12.5px}
.ov-split{display:grid;grid-template-columns:1fr 1fr;gap:22px;margin-top:14px;padding-top:12px;border-top:1px solid var(--line-2);position:relative}
@media (max-width:760px){.ov-split{grid-template-columns:1fr;gap:14px}}
.ov-side{min-width:0}
.ov-colh{display:flex;align-items:center;gap:7px;font-size:10.5px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px}
.ov-colh.ok{color:var(--ok)}.ov-colh.miss{color:var(--miss)}
.ov-colh .n{font-size:10px;font-weight:700;padding:1px 7px;border-radius:999px;background:currentColor;color:#fff;letter-spacing:0}
.ov-colh.ok .n{background:var(--ok)}.ov-colh.miss .n{background:var(--miss)}
.ov-mix{font-size:10.5px;font-weight:500;color:var(--ink-3);text-transform:none;letter-spacing:0}
.ov-mrow{display:grid;grid-template-columns:auto minmax(0,1fr) auto;column-gap:9px;row-gap:1px;align-items:center;padding:6px 0;border-bottom:1px dashed var(--line-2)}
.ov-stag{font-size:9px;font-weight:700;letter-spacing:.03em;padding:2px 0;border-radius:5px;width:34px;text-align:center;flex:none}
.ov-mname{font-size:12.5px;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ov-mval{font-size:11px;font-weight:700;padding:2px 9px;border-radius:999px;white-space:nowrap;max-width:150px;overflow:hidden;text-overflow:ellipsis;background:#fff;border:1px solid}
.ov-mval.ok{color:var(--ok);border-color:rgba(31,157,107,.32)}.ov-mval.warn{color:#b07012;border-color:rgba(214,150,40,.42)}.ov-mval.miss{color:var(--miss);border-color:rgba(207,61,61,.32)}
.ov-why{grid-column:2 / 4;font-size:11px;color:var(--ink-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ov-more{margin-top:8px}
.ov-none{font-size:12px;color:var(--ink-3);padding:6px 0}.ov-none.ok{color:var(--ok)}
.ov-dd-btn{display:inline-flex;align-items:center;gap:7px;font-family:inherit;font-size:11.5px;font-weight:600;color:var(--ink-2);background:#fff;border:1px solid var(--line);border-radius:999px;padding:4px 11px 4px 13px;cursor:pointer;box-shadow:0 1px 2px rgba(22,35,59,.05);transition:border-color .15s var(--ease),color .15s var(--ease)}
.ov-dd-btn:hover,.ov-dd-btn.on{border-color:var(--brand2);color:var(--steel)}
.ov-dd-btn svg{transition:transform .18s var(--ease)}.ov-dd-btn.on svg{transform:rotate(180deg)}
.ov-dd{position:fixed;z-index:40;display:flex;flex-direction:column;font-family:'Poppins',system-ui,sans-serif;color:var(--ink);background:#fff;border:1px solid var(--line);border-radius:14px;box-shadow:0 18px 44px -14px rgba(22,35,59,.32),0 2px 6px rgba(22,35,59,.06);overflow:hidden;animation:riseIn .16s var(--ease) both}
.ov-dd-h{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:9px 10px 9px 14px;border-bottom:1px solid var(--line-2);font-size:10.5px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;flex:none}
.ov-dd-h.ok{color:var(--ok)}.ov-dd-h.miss{color:var(--miss)}
.ov-dd-x{border:none;background:none;font-size:18px;line-height:1;color:var(--ink-3);cursor:pointer;padding:0 4px;font-family:inherit}.ov-dd-x:hover{color:var(--ink)}
.ov-dd-body{overflow:auto;padding:4px 14px 10px}
.ov-dd-grp{padding-top:8px}
.ov-dd-gh{display:flex;align-items:center;gap:8px;font-size:11px;font-weight:600;color:var(--ink-2);padding-bottom:3px}
.ov-dd-gn{margin-left:auto;font-size:10.5px;color:var(--ink-3);font-weight:600}
.ov-dd-row{display:grid;grid-template-columns:minmax(0,1fr) auto;column-gap:10px;row-gap:1px;align-items:center;padding:5px 0 5px 43px;border-bottom:1px dashed var(--line-2)}
.ov-dd-row:last-child{border-bottom:none}
.ov-dd-row .ov-mname{font-size:12px}
.ov-dd-row .ov-why{grid-column:1 / 3}
.view-switch .segbtn{padding:7px 18px}
.hist-ctl{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px}
.hist-ctl select{font-size:12.5px;padding:7px 10px;border-radius:10px;border:1px solid var(--line);background:var(--surface);font-family:inherit}
.hist-note{font-size:11.5px;color:var(--ink-3)}
.hist-wrap{overflow-x:auto;border:1px solid var(--line);border-radius:14px;background:var(--surface);box-shadow:var(--sh-1)}
table.hist{border-collapse:separate;border-spacing:0;width:100%;font-size:12.5px}
table.hist th,table.hist td{padding:8px 11px;border-bottom:1px solid var(--line-2);text-align:center;white-space:nowrap}
table.hist th{font-size:11px;font-weight:600;color:var(--ink-2);background:#f7f9fb;vertical-align:bottom}
table.hist .hm{position:sticky;left:0;z-index:1;background:var(--surface);text-align:left;min-width:200px;max-width:260px;white-space:normal;border-right:1px solid var(--line-2)}
table.hist th.hm{background:#f7f9fb;z-index:2}
.hist-name{display:block;font-weight:500;color:var(--ink)}
.hist-target{display:block;font-size:10.5px;color:var(--ink-3);margin-top:1px}
.hist-c{font-variant-numeric:tabular-nums;font-weight:500;min-width:76px}
table.hist td.hist-c{white-space:normal;max-width:190px;line-height:1.35}
.hist-c.ok{background:var(--ok-bg);color:#1c6b4b}.hist-c.warn{background:var(--warn-bg);color:#7a5209}.hist-c.miss{background:var(--miss-bg);color:#8a2b2b}
.hist-c.none{color:var(--ink-3)}
table.hist td.gap,table.hist th.gap{background:repeating-linear-gradient(135deg,#f7f8fa 0 6px,#eef1f4 6px 12px)}
.hist-miss{display:block;font-size:9px;font-weight:700;color:var(--ink-3);text-transform:uppercase;letter-spacing:.05em;margin-top:2px}
.hist-g td{background:#fbfcfd;font-size:11.5px;font-weight:600;color:var(--ink-2);padding-top:13px}
.hist-g td.hm{background:#fbfcfd}
.hist-r{cursor:pointer}.hist-r:hover td.hm{background:var(--steel-soft)}
.hist-r.on td.hm{box-shadow:inset 3px 0 0 var(--steel);background:var(--steel-soft)}
.hist-overall td{font-weight:600}.hist-overall .hist-name{font-weight:700}
.hist-chg,.hist-chg-h{min-width:92px}
.hist-chg{font-size:11.5px;font-weight:600}.hist-chg.up{color:var(--ok)}.hist-chg.down{color:var(--miss)}.hist-chg.flat{color:var(--ink-3)}
.hist-chart-row td{background:#fbfcfd;text-align:left;white-space:normal;padding:12px 14px}
.hist-ph{border:none;background:none;font:inherit;color:inherit;cursor:pointer;padding:0}
.hist-ph:hover{color:var(--steel);text-decoration:underline}
.trend{position:sticky;left:14px}
.trend-h{display:flex;gap:10px;align-items:baseline;font-size:12.5px;margin-bottom:6px}
.trend-h span{color:var(--ink-3);font-size:11.5px}
.trend-none{font-size:12px;color:var(--ink-3)}
.trend-tip{position:absolute;pointer-events:none;background:var(--surface);border:1px solid var(--line);border-radius:10px;box-shadow:var(--sh-2);padding:7px 11px;font-size:11.5px;white-space:nowrap;transform:translateX(-50%);z-index:3}
.trend-tip b{display:block;font-size:13.5px;color:var(--ink)}
.trend-tip span{display:block;color:var(--ink-3)}
.trend-tip .trend-why{color:var(--ink-2);max-width:260px;white-space:normal}
.ov-ctl{gap:10px;flex-wrap:wrap;justify-content:flex-end}
.ov-step{display:flex;align-items:center;gap:4px}
.ov-arrow{font-size:20px;padding:4px 10px;border:1px solid var(--line);background:var(--surface);border-radius:10px}
.ov-streak{display:inline-flex;align-items:center;gap:7px;margin-top:9px;font-size:11.5px;color:#8a2b2b;background:#fff;border:1px solid rgba(207,61,61,.3);border-radius:999px;padding:3px 11px 3px 4px}
.ov-streak b{font-weight:600}
.ov-streak-ic{width:17px;height:17px;border-radius:50%;background:var(--miss);color:#fff;font-size:11px;font-weight:700;display:inline-flex;align-items:center;justify-content:center}
.arch{margin-top:16px}.arch-note{font-size:11.5px;color:var(--ink-3);padding:10px 14px}
.ownertag{font-size:11.5px;font-weight:600;padding:2px 9px;border-radius:20px;background:#eef1f5;color:var(--ink-2);white-space:nowrap}
.msg.info{background:var(--steel-soft);color:var(--steel)}
.people-h{font-size:14px;font-weight:600;color:var(--navy);margin:22px 0 10px}
.inline-select{font-size:12.5px;padding:5px 8px;border-radius:8px;border:1px solid var(--line);background:var(--surface);font-family:inherit;max-width:190px}
.msg{font-size:12.5px;padding:9px 12px;border-radius:8px;margin-bottom:12px}.msg.err{background:var(--miss-bg);color:var(--miss)}.msg.ok{background:var(--ok-bg);color:var(--ok)}
.note-box{background:var(--warn-bg);color:#7a5209;font-size:12px;border-radius:9px;padding:11px 14px;line-height:1.45;margin-bottom:16px}
.chips{display:flex;gap:6px;flex-wrap:wrap}
.pill{font-size:11.5px;font-weight:500;padding:4px 11px;border-radius:20px;background:var(--steel-soft);color:var(--steel)}
.reportgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(238px,1fr));gap:14px;margin-bottom:6px}
.report-card{background:var(--surface);border:1px solid var(--line);border-radius:18px;padding:17px 16px 15px;box-shadow:var(--sh-1);position:relative;overflow:hidden;transition:transform .16s var(--ease),box-shadow .18s var(--ease)}
.report-card::before{content:"";position:absolute;top:0;left:0;right:0;height:3px;background-image:linear-gradient(90deg,transparent,rgba(255,255,255,.95) 50%,transparent),linear-gradient(90deg,var(--brand1),var(--brand2),var(--brand3));background-size:28% 100%,100% 100%;background-repeat:no-repeat;background-position:-60% 0,0 0;animation:streak 3.4s linear infinite}
.report-card:hover{transform:translateY(-2px);box-shadow:var(--sh-2)}
.report-card.dragover{border-color:var(--steel);background:var(--steel-soft);box-shadow:0 0 0 3px rgba(15,159,136,.18),var(--sh-2)}
.rc-top{display:flex;justify-content:space-between;align-items:flex-start;gap:8px}
.rc-name{font-weight:600;font-size:13.5px}
.rc-hint{font-size:11.5px;color:var(--ink-3);margin-top:3px;line-height:1.35}
.rc-figs{display:flex;flex-direction:column;gap:6px;margin:12px 0 8px;padding-top:10px;border-top:1px solid var(--line-2)}
.rc-fig{display:flex;justify-content:space-between;align-items:baseline;font-size:12.5px;gap:8px}
.rc-fig .k{color:var(--ink-2)}
.rc-fig .v{font-weight:600}
.rc-fig .v.miss{color:var(--miss);font-weight:500;font-size:11.5px}
.rc-file{font-size:10.5px;color:var(--ink-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.calpop{position:absolute;top:calc(100% + 6px);left:0;z-index:30;background:var(--surface);border:1px solid var(--line);border-radius:12px;box-shadow:0 16px 40px rgba(24,28,34,.18);padding:12px;width:262px}
.calhead{display:flex;align-items:center;justify-content:space-between;font-weight:600;font-size:13.5px;margin-bottom:8px}
.calnav{border:none;background:none;font-size:18px;line-height:1;cursor:pointer;color:var(--ink-2);padding:2px 9px;border-radius:6px}
.calnav:hover{background:var(--line-2)}
.calgrid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px}
.calgrid.caldow{margin-bottom:4px}
.dow{text-align:center;font-size:10.5px;color:var(--ink-3);padding:2px 0}
.calday{border:none;background:none;font-family:inherit;font-size:12.5px;padding:7px 0;border-radius:6px;cursor:pointer;color:var(--ink)}
.calday:hover{background:var(--steel-soft)}
.calday.in{background:var(--steel-soft);color:var(--steel)}
.calday.edge{background:var(--steel);color:#fff;font-weight:600}
.pp{position:relative;display:inline-block}
.pp-btn{min-width:210px;display:flex;align-items:center;justify-content:space-between;gap:10px}
.pp-btn.on{border-color:var(--brand2)}
.pp-pop{width:300px}.pp-pop.right{left:auto;right:0}
.pp-tabs{display:flex;width:100%;margin-bottom:10px}.pp-tabs .segbtn{flex:1;padding:6px 4px;font-size:12px}
.pp-months{display:grid;grid-template-columns:repeat(4,1fr);gap:5px}
.pp-month{position:relative;border:1px solid var(--line-2);background:none;font-family:inherit;font-size:12.5px;padding:10px 0;border-radius:8px;cursor:pointer;color:var(--ink)}
.pp-month:hover{background:var(--steel-soft)}.pp-month.on{background:var(--steel);border-color:var(--steel);color:#fff;font-weight:600}
.pp-dot{position:absolute;top:4px;right:5px;width:6px;height:6px;border-radius:50%;background:var(--brand2)}
.pp-month.on .pp-dot{background:#fff}
.pp-weeks{display:flex;flex-direction:column;gap:2px}
.pp-week{position:relative;display:grid;grid-template-columns:repeat(7,1fr);gap:2px;border-radius:8px}
.pp-week.pick{cursor:pointer}.pp-week.pick:hover{background:var(--steel-soft)}
.pp-week.pick .calday{pointer-events:none}
.pp-week.on{background:var(--steel)}.pp-week.on .calday{color:#fff;font-weight:600}
.pp-week .pp-dot.wk{top:50%;right:-9px;margin-top:-3px}
.calday{position:relative}.calday .pp-dot{top:3px;right:3px;width:5px;height:5px}
.calday.out{color:var(--ink-3)}
.pp-hint{font-size:11px;color:var(--ink-3);margin-top:8px;line-height:1.45}
.pp-card .pp-lab{font-size:12.5px;font-weight:600;color:var(--ink-2)}
.pp-type{font-size:10.5px;font-weight:600;padding:3px 9px;border-radius:999px;background:var(--steel-soft);color:var(--steel)}
.pp-recent{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-top:12px;padding-top:12px;border-top:1px solid var(--line-2)}
.pp-recent .pp-lab{font-size:11px;font-weight:600;color:var(--ink-3);text-transform:uppercase;letter-spacing:.05em;margin-right:2px}
.pp-chip{border:none;cursor:pointer;font-family:inherit;font-size:11.5px;font-weight:500;padding:5px 11px;border-radius:999px;background:var(--steel-soft);color:var(--steel)}
.pp-chip.on{background:var(--steel);color:#fff}
.pp-all{font-size:12px;padding:5px 8px;border-radius:999px;border:1px solid var(--line);background:var(--surface);font-family:inherit;color:var(--ink-2);max-width:220px}
.calfoot{display:flex;align-items:center;justify-content:space-between;margin-top:10px;padding-top:10px;border-top:1px solid var(--line-2)}
.btn.primary.sm{padding:6px 12px;font-size:12px}
.scopework{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px;margin-bottom:6px}
.scopework-card{background:var(--surface);border:1px solid var(--line);border-left:3px solid var(--brand2);border-radius:16px;padding:13px 15px;box-shadow:var(--sh-1);transition:transform .16s var(--ease),box-shadow .18s var(--ease)}
.scopework-card:hover{transform:translateY(-2px);box-shadow:var(--sh-2);border-left-color:var(--brand3)}
.sw-name{font-size:13px;font-weight:600}
.sw-sub{font-size:11.5px;color:var(--ink-3);margin-top:3px}
.volgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:8px}
.volcard{background:var(--surface);border:1px solid var(--line);border-radius:18px;padding:14px 16px;box-shadow:var(--sh-1);transition:transform .16s var(--ease),box-shadow .18s var(--ease)}
.volcard:hover{transform:translateY(-2px);box-shadow:var(--sh-2)}
.vl{font-size:11.5px;color:var(--ink-3)}
.vv{font-size:21px;font-weight:700;margin-top:4px;letter-spacing:-.01em}
.vd{font-size:11px;margin-top:3px;color:var(--ink-3)}
.bulkzone{position:relative;overflow:hidden;display:flex;gap:16px;align-items:flex-start;border-radius:20px;padding:18px 20px;margin:0 0 16px;border:1.5px dashed rgba(44,201,176,.45);background:radial-gradient(130% 160% at 0% 0%,rgba(62,207,124,.1),transparent 58%),linear-gradient(140deg,#fff 52%,rgba(56,200,232,.08));box-shadow:var(--sh-1);transition:border-color .16s var(--ease),box-shadow .16s var(--ease),transform .16s var(--ease)}
.bulkzone::before{content:"";position:absolute;top:0;left:0;right:0;height:3px;background-image:linear-gradient(90deg,transparent,rgba(255,255,255,.95) 50%,transparent),var(--grad);background-size:60% 100%,100% 100%;background-repeat:no-repeat,no-repeat;animation:streak 3.4s linear infinite}
.bulkzone.on{border-color:var(--brand2);box-shadow:var(--sh-3);transform:translateY(-1px)}
.bulkzone>*{position:relative;z-index:1}
.bz-icon{width:42px;height:42px;flex:none;border-radius:13px;display:flex;align-items:center;justify-content:center;font-size:20px;color:#fff;background:var(--grad);box-shadow:0 8px 18px -8px rgba(44,201,176,.8)}
.bz-title{font-size:14px;font-weight:600}
.bz-sub{font-size:11.5px;color:var(--ink-3);margin-top:3px;max-width:62ch;line-height:1.5}
.bz-count{font-size:11.5px;color:var(--ink-3);align-self:center}
.rc-err{font-size:11px;color:var(--miss);background:var(--miss-bg);border-radius:9px;padding:6px 9px;margin-top:9px}
.modal.wide{max-width:760px;width:94vw}
.bulktable{max-height:46vh;overflow:auto;border:1px solid var(--line);border-radius:14px;margin:4px 0 10px}
.bulkrow{display:flex;align-items:center;gap:10px;padding:9px 12px;border-bottom:1px solid var(--line-2)}
.bulkrow:last-child{border-bottom:none}
.bulkrow.dup{background:var(--miss-bg)}
.br-file{flex:1;min-width:0}
.br-name{font-size:12.5px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.br-reason{font-size:10.5px;color:var(--ink-3);margin-top:2px}
.br-sel{font-family:inherit;font-size:11.5px;padding:6px 8px;border-radius:9px;border:1px solid var(--line);background:#fff;color:var(--ink);max-width:230px}
.bulkfoot{font-size:12px;color:var(--ink-2);margin-bottom:4px}
.nav{flex-wrap:nowrap}.nav .tab{white-space:nowrap;flex:none}.who{flex:none;white-space:nowrap}.who-mail{max-width:240px;overflow:hidden;text-overflow:ellipsis}@media(max-width:1240px){.who-mail{display:none}.nav .tab{padding:9px 11px}.nav .brand{margin-right:8px}}@media(max-width:900px){.nav{overflow-x:auto;scrollbar-width:none}.nav::-webkit-scrollbar{display:none}.nav .bytag{display:none}}
.fm-tools{display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:14px}
.fm-search{flex:1;min-width:230px;max-width:360px;font-family:inherit;font-size:13px;padding:10px 14px;border-radius:999px;border:1px solid var(--line);background:#fff;color:var(--ink);outline:none}
.fm-search:focus{border-color:var(--brand2);box-shadow:0 0 0 3px rgba(44,201,176,.16)}
.fm-chips{display:flex;gap:6px;flex-wrap:wrap}
.fm-chip{font-family:inherit;font-size:11.5px;font-weight:500;padding:6px 12px;border-radius:999px;border:1px solid var(--line);background:#fff;color:var(--ink-2);cursor:pointer;transition:all .15s var(--ease)}
.fm-chip:hover{border-color:var(--brand2)}
.fm-chip.on{background:var(--steel);border-color:var(--steel);color:#fff}
.fm-how{padding:18px 20px;margin-bottom:14px}
.fm-how-h{font-size:14px;font-weight:600;margin-bottom:10px}
.fm-how-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
.fm-how-grid>div{background:rgba(255,255,255,.7);border:1px solid var(--line);border-radius:14px;padding:11px 13px}
.fm-how-grid b{font-size:12px;font-weight:600}
.fm-how-grid p{margin:4px 0 0;font-size:11.5px;line-height:1.5;color:var(--ink-2)}
.fm-note{font-size:11.5px;color:var(--ink-3);margin-top:11px;line-height:1.5}
.fm-scope{padding:18px 20px;margin-bottom:14px}
.fm-scope-h{display:flex;gap:12px;align-items:center;margin-bottom:12px}
.fm-badge{width:40px;height:40px;flex:none;border-radius:12px;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:10.5px}
.fm-scope-name{font-size:15px;font-weight:600}
.fm-scope-sub{font-size:11.5px;color:var(--ink-3);margin-top:1px}
.fm-table{border:1px solid var(--line);border-radius:14px;overflow:hidden;background:rgba(255,255,255,.75)}
.fm-row{display:grid;grid-template-columns:1.05fr 1.7fr 1.25fr 1.1fr;gap:14px;padding:11px 14px;border-bottom:1px solid var(--line-2);align-items:start}
.fm-row:last-child{border-bottom:none}
.fm-head{font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-3);background:var(--line-2);padding:8px 14px}
.fm-kpi{font-size:13px;font-weight:600;display:flex;flex-direction:column;gap:5px;align-items:flex-start}
.fm-tag{font-size:9.5px;font-weight:600;padding:2px 8px;border-radius:999px}
.fm-tag.auto{background:var(--steel-soft);color:var(--steel)}
.fm-tag.man{background:var(--info-bg);color:var(--info)}
.fm-formula{display:inline-block;font-size:12.5px;line-height:1.5;background:var(--line-2);border-radius:9px;padding:6px 10px;color:var(--ink)}
.fm-alt{font-size:11px;color:var(--ink-3);margin-top:6px;line-height:1.45}
.fm-plain{font-size:12px;color:var(--ink-3);line-height:1.5}
.fm-bands{display:flex;flex-direction:column;gap:5px;align-items:flex-start}
.fm-bands .chip,.fm-sla-f .chip{font-size:10.5px;font-weight:600;padding:3px 9px;border-radius:999px}
.fm-from{font-size:12px;color:var(--ink-2);line-height:1.5}
.fm-sla{margin-top:12px;border:1px dashed var(--line);border-radius:14px;padding:12px 14px}
.fm-sla-h{font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-3);margin-bottom:8px}
.fm-sla-list{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px}
.fm-sla-item{font-size:11.5px;padding:4px 10px;border-radius:999px;background:#fff;border:1px solid var(--line);color:var(--ink-2)}
.fm-sla-f{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
@media(max-width:980px){.fm-how-grid{grid-template-columns:repeat(2,1fr)}.fm-row{grid-template-columns:1fr;gap:7px}.fm-head{display:none}.fm-bands{flex-direction:row;flex-wrap:wrap}}
@media(max-width:560px){.fm-how-grid{grid-template-columns:1fr}}
.bfp-list{display:flex;flex-direction:column;gap:8px;margin:6px 0 4px}
.bfp-item{display:flex;justify-content:space-between;align-items:center;gap:12px;font-family:inherit;text-align:left;padding:12px 14px;border-radius:14px;border:1px solid var(--line);background:#fff;cursor:pointer;transition:border-color .15s var(--ease),box-shadow .15s var(--ease)}
.bfp-item:hover{border-color:var(--brand2);box-shadow:var(--sh-2)}
.bfp-name{font-size:13.5px;font-weight:600;color:var(--ink)}
.bfp-count{font-size:11.5px;color:var(--ink-3);white-space:nowrap}
.od-bar{margin:0 0 14px;padding:15px 18px;border-radius:18px;background:#fff;border:1.5px solid rgba(44,201,176,.55);box-shadow:var(--sh-2)}
.od-top{display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.od-ic{width:38px;height:38px;border-radius:11px;background:linear-gradient(135deg,var(--brand1),var(--brand3));display:flex;align-items:center;justify-content:center;flex:none}
.od-main{flex:1;min-width:260px}
.od-title{font-weight:600;font-size:14.5px;display:flex;align-items:center;gap:8px}
.ownertag.od-on{background:var(--ok-bg);color:var(--ok)}
.od-sub{font-size:12px;color:var(--ink-3);margin-top:3px;line-height:1.5}
.od-acct{font-size:12px;color:var(--ink-2);text-align:right;line-height:1.5}
.od-muted{color:var(--ink-3);font-size:12.5px}
.od-actions{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-top:13px;padding-top:13px;border-top:1px solid var(--line-2)}
.od-status{font-size:12px;color:var(--ink-3);line-height:1.5;min-width:0;flex:1 1 320px}
.od-status a.link{white-space:nowrap}
.od-err{margin-top:12px;font-size:12px;padding:8px 11px}
.od-done{margin-top:10px;font-size:12px;color:var(--ok);font-weight:500}
.od-hint{font-size:11.5px;color:var(--ink-3);margin:-4px 0 8px 4px}
.od-bar.od-off{display:flex;align-items:center;gap:11px;padding:10px 14px;border-color:var(--line);box-shadow:none}
.od-ic.sm{width:28px;height:28px;border-radius:9px}
.od-ic.sm svg{width:16px;height:12px}
.od-off-t{font-size:12.5px;color:var(--ink-2)}
.od-steps{margin:0 0 8px;padding-left:20px;font-size:12.5px;color:var(--ink-2);line-height:1.7}
.od-note{font-size:12px;color:var(--ok);font-weight:500;margin:0 0 14px}
.od-crumbs{display:flex;align-items:center;gap:7px;flex-wrap:wrap;font-size:12.5px;margin:-4px 0 10px;color:var(--ink-2)}
.od-sep{color:var(--ink-3)}
.od-review-src{font-size:12px;color:var(--ink-3);margin:-8px 0 12px}
.od-browser .bfp-list{max-height:52vh;overflow:auto;padding-right:2px}
.od-file{display:flex;align-items:center;gap:11px;padding:10px 14px;border-radius:14px;border:1px solid var(--line);background:#fff;cursor:pointer;font-size:13.5px;color:var(--ink)}
.od-file:hover{border-color:rgba(44,201,176,.55)}
.od-file input{width:16px;height:16px;flex:none;accent-color:var(--brand2);cursor:pointer}
.od-file-n{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.od-file-d{font-size:11.5px;color:var(--ink-3);white-space:nowrap}
.od-all{border-style:dashed;background:transparent;font-weight:600}
a.link{display:inline}
.gd-layout{display:grid;grid-template-columns:220px minmax(0,1fr);gap:18px;align-items:start}
.gd-layout nav.gd-toc.panel{position:sticky;top:78px;overflow:auto}.gd-toc{padding:14px 10px;display:flex;flex-direction:column;gap:1px;max-height:calc(100vh - 96px);overflow:auto}
.gd-toc-h{font-size:10px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;color:var(--ink-3);padding:0 8px 8px}
.gd-toc button{font-family:inherit;font-size:12px;text-align:left;border:none;background:none;color:var(--ink-2);padding:6px 8px;border-radius:9px;cursor:pointer;line-height:1.35}
.gd-toc button:hover{background:var(--steel-soft);color:var(--steel)}
.gd-sec{padding:20px 22px;margin-bottom:14px;scroll-margin-top:76px}
.gd-h{display:flex;gap:13px;align-items:flex-start;margin-bottom:10px}
.gd-h h2{font-size:17px;font-weight:600;margin:0;letter-spacing:-.01em}
.gd-num{width:34px;height:34px;flex:none;border-radius:11px;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:15px;color:#fff;background:var(--grad);box-shadow:0 8px 16px -8px rgba(44,201,176,.8)}
.gd-lead{margin:4px 0 0;font-size:13px;color:var(--ink-2);line-height:1.55}
.gd-sec h3{font-size:13px;font-weight:600;margin:18px 0 8px;color:var(--ink)}
.gd-sec p{font-size:13px;line-height:1.65;color:var(--ink-2);margin:8px 0}
.gd-sec em{font-style:normal;font-weight:600;color:var(--ink)}
.gd-steps{margin:8px 0;padding:0;list-style:none;counter-reset:gd;display:flex;flex-direction:column;gap:8px}
.gd-steps>li{counter-increment:gd;position:relative;padding:9px 12px 9px 44px;font-size:13px;line-height:1.6;color:var(--ink-2);background:rgba(255,255,255,.72);border:1px solid var(--line);border-radius:13px}
.gd-steps>li::before{content:counter(gd);position:absolute;left:11px;top:9px;width:22px;height:22px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:var(--steel);background:var(--steel-soft)}
.gd-steps b,.gd-list b{color:var(--ink);font-weight:600}
.gd-list{margin:8px 0;padding-left:18px;display:flex;flex-direction:column;gap:6px}
.gd-list li{font-size:13px;line-height:1.6;color:var(--ink-2)}
.gd-note{margin:12px 0 4px;padding:11px 14px;border-radius:13px;font-size:12.5px;line-height:1.6;color:var(--ink-2);background:var(--steel-soft);border:1px solid rgba(15,159,136,.18)}
.gd-note>b:first-child{display:block;font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--steel);margin-bottom:3px}.gd-note b{font-weight:600;color:var(--ink)}
.gd-note.warn{background:var(--warn-bg);border-color:rgba(192,125,16,.22)}
.gd-note.warn>b:first-child{color:var(--warn)}
.gd-note em{font-style:normal;font-weight:600;color:var(--ink)}
.gd-table{border:1px solid var(--line);border-radius:13px;overflow:hidden;margin:10px 0;background:rgba(255,255,255,.75)}
.gd-tr{display:grid;grid-template-columns:minmax(150px,.42fr) 1fr;gap:14px;padding:9px 13px;border-bottom:1px solid var(--line-2);font-size:12.5px;line-height:1.55;color:var(--ink-2);align-items:start}
.gd-tr:last-child{border-bottom:none}
.gd-tr b{color:var(--ink);font-weight:600}
.gd-th{font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-3);background:var(--line-2);padding:7px 13px}
.gd-tr .chip{font-size:10.5px;font-weight:600;padding:3px 9px;border-radius:999px}
.gd-flow{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:12px 0}
.gd-flow span{font-size:12px;font-weight:600;padding:7px 13px;border-radius:999px;background:#fff;border:1px solid var(--line);color:var(--ink)}
.gd-flow i{font-style:normal;color:var(--brand2);font-weight:700}
.gd-scope{margin-top:12px}
.gd-scope-h{display:flex;align-items:center;gap:9px;font-size:13px;font-weight:600}
.gd-you{font-size:12px!important;color:var(--ink-3)!important}
@media(max-width:980px){.gd-layout{grid-template-columns:1fr}.gd-toc{position:static;flex-direction:row;flex-wrap:wrap;max-height:none}.gd-toc-h{width:100%}.gd-tr{grid-template-columns:1fr;gap:3px}}
.cp-why{font-size:11.5px;color:var(--ink-3);margin:-2px 0 6px;line-height:1.5}
.rc-cols{margin-top:9px;padding-top:8px;border-top:1px dashed var(--line)}
.rc-cols-h{font-size:9.5px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-3);margin-bottom:4px}
.rc-col{display:flex;justify-content:space-between;gap:10px;font-size:11px;line-height:1.7;color:var(--ink-3)}
.rc-col b{font-weight:600;color:var(--ink-2);text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:60%}
.rc-saved{font-size:10.5px;color:var(--ink-3);margin-top:6px}
.calcrow{grid-column:1/-1;display:flex;gap:10px;flex-wrap:wrap;align-items:center;font-size:11.5px;color:var(--ink-3);background:var(--info-bg);border-radius:10px;padding:6px 11px;margin:2px 0 6px}.calcrow b{color:var(--ink);font-weight:600}
.confetti-wrap{position:fixed;inset:0;pointer-events:none;z-index:100;overflow:hidden}
.confetti-pc{position:absolute;top:-14px;border-radius:2px;animation-name:confetti-fall;animation-timing-function:ease-in;animation-fill-mode:forwards}
@keyframes confetti-fall{0%{transform:translateY(0) rotate(0);opacity:1}100%{transform:translateY(106vh) rotate(600deg);opacity:.85}}
/* ---- floating feedback widget ---- */
.fb-tab{position:fixed;z-index:60;width:34px;min-height:132px;border:none;cursor:grab;color:#fff;font-family:inherit;font-size:12.5px;font-weight:600;letter-spacing:.04em;border-radius:12px 0 0 12px;background:linear-gradient(160deg,var(--brand1),var(--brand2),var(--brand3));box-shadow:-4px 6px 20px -6px rgba(20,120,110,.55);display:flex;align-items:center;justify-content:center;padding:14px 0;touch-action:none;user-select:none;transition:box-shadow .18s var(--ease),filter .18s var(--ease)}
.fb-tab:hover{filter:brightness(1.05);box-shadow:-6px 8px 26px -6px rgba(20,120,110,.7)}
.fb-tab:active{cursor:grabbing}
.fb-tab .fb-lab{writing-mode:vertical-rl;transform:rotate(180deg);line-height:1}
.fb-tab::before{content:"";position:absolute;top:0;left:0;height:34%;width:3px;border-radius:2px;background:linear-gradient(180deg,rgba(255,255,255,.9),transparent);animation:lineStreak 3.2s linear infinite}
.fb-panel{position:fixed;z-index:61;width:320px;background:var(--surface);border:1px solid var(--line);border-radius:18px;box-shadow:0 26px 70px -22px rgba(16,26,42,.5);overflow:hidden;animation:riseIn .22s var(--ease) both}
.fb-panel::before{content:"";position:absolute;top:0;left:0;right:0;height:3px;background-image:linear-gradient(90deg,transparent,rgba(255,255,255,.95) 50%,transparent),linear-gradient(90deg,var(--brand1),var(--brand2),var(--brand3));background-size:28% 100%,100% 100%;background-repeat:no-repeat;background-position:-60% 0,0 0;animation:streak 3.4s linear infinite}
.fb-head{padding:16px 18px 10px}
.fb-head .t{font-size:15px;font-weight:700;letter-spacing:-.01em}
.fb-head .s{font-size:11.5px;color:var(--ink-3);margin-top:2px}
.fb-body{padding:4px 18px 18px}
.fb-kinds{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px}
.fb-k{border:1px solid var(--line);background:var(--surface);color:var(--ink-2);font-family:inherit;font-size:11.5px;font-weight:500;padding:7px 11px;border-radius:999px;cursor:pointer;transition:all .14s var(--ease)}
.fb-k:hover{border-color:var(--brand2)}
.fb-k.on{color:#fff;border-color:transparent;background:linear-gradient(135deg,var(--brand1),var(--brand3));font-weight:600}
.fb-ta{width:100%;min-height:96px;resize:vertical;font-size:13px}
.fb-meta{font-size:10.5px;color:var(--ink-3);line-height:1.45;margin:8px 0 12px}
.fb-foot{display:flex;align-items:center;justify-content:space-between;gap:10px}
.fb-count{font-size:11px;color:var(--ink-3)}
.fb-thanks{padding:26px 20px;text-align:center}
.fb-thanks .ic{width:46px;height:46px;border-radius:50%;margin:0 auto 12px;background:var(--ok-bg);color:var(--ok);display:flex;align-items:center;justify-content:center;font-size:22px}
.fb-x{position:absolute;top:10px;right:12px;border:none;background:none;color:var(--ink-3);font-size:18px;cursor:pointer;line-height:1;padding:4px;border-radius:6px}
.fb-x:hover{background:var(--line-2);color:var(--ink)}
.fbadm-row{border:1px solid var(--line);border-left-width:3px;border-radius:12px;padding:12px 14px;margin-bottom:10px;background:var(--surface)}
.fbadm-row.bug{border-left-color:var(--miss)}.fbadm-row.idea{border-left-color:var(--warn)}.fbadm-row.praise{border-left-color:var(--ok)}
/* ---- shout-out / recognition ---- */
.shout{position:relative;overflow:hidden;border-radius:20px;padding:16px 20px;margin-bottom:14px;color:#4a3206;background:linear-gradient(135deg,#f9dd86,#eab63f 55%,#f6d365);box-shadow:0 12px 34px -14px rgba(200,150,30,.6)}
.shout::after{content:"";position:absolute;top:0;left:0;height:100%;width:34%;background:linear-gradient(100deg,transparent,rgba(255,255,255,.6),transparent);transform:translateX(-160%) skewX(-16deg);animation:sheen 4.5s ease-in-out infinite;pointer-events:none}
.shout>*{position:relative;z-index:1}
.shout .kbtn{font-family:inherit;font-size:12px;font-weight:700;padding:8px 15px;border-radius:999px;border:none;cursor:pointer;background:rgba(255,255,255,.55);color:#5a3d09;white-space:nowrap;transition:background .16s var(--ease)}
.shout .kbtn:hover{background:rgba(255,255,255,.8)}
.ov-card.star{border-color:rgba(230,169,58,.55)}
.starribbon{position:absolute;top:13px;right:-36px;transform:rotate(38deg);width:150px;text-align:center;font-size:9.5px;font-weight:700;letter-spacing:.06em;color:#5a3d09;padding:4px 0;background:linear-gradient(135deg,#f6d365,#e6a93a,#f4cd6a);box-shadow:0 2px 8px -2px rgba(180,130,20,.6);z-index:3}
.bestchip{font-size:10.5px;font-weight:600;padding:3px 10px;border-radius:999px;background:linear-gradient(135deg,#fbeecb,#f6dd9b);color:#8a5a10;border:1px solid rgba(230,169,58,.4)}
@media(max-width:760px){.crow{grid-template-columns:1fr 1fr;gap:8px}.crow .cres{justify-content:flex-start}}
/* ---- shine / motion ---- */
@keyframes sheen{0%{transform:translateX(-140%) skewX(-16deg)}55%,100%{transform:translateX(360%) skewX(-16deg)}}
@keyframes gradpan{0%{background-position:0% 50%}100%{background-position:200% 50%}}
@keyframes lineflow{0%{background-position:-200% 0}100%{background-position:200% 0}}
@keyframes streak{0%{background-position:-60% 0,0 0}100%{background-position:160% 0,0 0}}
/* ---- travelling border beam + tinted surfaces for bordered sections ---- */
@property --bang{syntax:'<angle>';inherits:false;initial-value:0deg}
@keyframes borderSpin{to{--bang:360deg}}
.panel,.card,.scopework-card,.report-card,.scope-block,.volcard,.tile{position:relative;background:radial-gradient(rgba(44,201,176,.11) 1.2px,transparent 1.6px) 0 0/16px 16px,linear-gradient(135deg,#ffffff 0%,#eef9f5 55%,#e6f3f9 100%)}
.panel::after,.card::after,.scopework-card::after,.report-card::after,.scope-block::after,.volcard::after,.tile::after{content:"";position:absolute;inset:0;border-radius:inherit;padding:2px;background:conic-gradient(from var(--bang),transparent 0 70%,var(--brand1) 80%,rgba(255,255,255,.95) 88%,var(--brand3) 95%,transparent 100%);-webkit-mask:linear-gradient(#000 0 0) content-box,linear-gradient(#000 0 0);-webkit-mask-composite:xor;mask:linear-gradient(#000 0 0) content-box,linear-gradient(#000 0 0);mask-composite:exclude;animation:borderSpin 4.2s linear infinite;pointer-events:none;z-index:4}
.card::after{animation-duration:5s}.panel::after{animation-duration:6s}.scope-block::after{animation-duration:5.5s}
.tile.ok::after{background:conic-gradient(from var(--bang),transparent 0 70%,var(--ok) 80%,rgba(255,255,255,.95) 88%,var(--brand1) 95%,transparent 100%)}
.tile.warn::after{background:conic-gradient(from var(--bang),transparent 0 70%,var(--warn) 80%,rgba(255,255,255,.95) 88%,#f0c265 95%,transparent 100%)}
.tile.miss::after{background:conic-gradient(from var(--bang),transparent 0 70%,var(--miss) 80%,rgba(255,255,255,.95) 88%,#f0999a 95%,transparent 100%)}
.scope-head{background:linear-gradient(135deg,rgba(255,255,255,.75),rgba(238,241,245,.8))}
th{background:rgba(238,241,245,.75)}
@keyframes lineStreak{0%{transform:translateX(-120%)}100%{transform:translateX(560%)}}
@keyframes floaty{0%,100%{transform:translateY(0)}50%{transform:translateY(-3px)}}
@keyframes ringGlow{0%,100%{filter:drop-shadow(0 0 0 rgba(44,201,176,0))}50%{filter:drop-shadow(0 0 9px rgba(44,201,176,.45))}}
.wm-grad{animation:gradpan 7s linear infinite alternate}
.btn.primary{position:relative;overflow:hidden}
.btn.primary::after{content:"";position:absolute;top:0;left:0;height:100%;width:36%;background:linear-gradient(100deg,transparent,rgba(255,255,255,.55),transparent);transform:translateX(-160%) skewX(-16deg);pointer-events:none}
.btn.primary:hover::after{animation:sheen 1.1s ease-out}
/* ---- Overview: borderless, floating interface ---- */
.ov-hero{position:relative;overflow:hidden;border-radius:24px;padding:26px 30px;margin-bottom:16px;color:#eafaf5;background:linear-gradient(135deg,#0e3b38 0%,#114c45 52%,#0b6a62 100%);box-shadow:0 22px 54px -22px rgba(9,55,50,.7),0 2px 10px -4px rgba(22,35,59,.2)}
.ov-hero::before{content:"";position:absolute;inset:0;background:radial-gradient(70% 130% at 100% 0%,rgba(56,200,232,.28),transparent 55%),radial-gradient(70% 140% at 0% 120%,rgba(62,207,124,.24),transparent 55%);pointer-events:none}
.ov-hero::after{content:"";position:absolute;top:0;left:0;height:100%;width:40%;background:linear-gradient(100deg,transparent,rgba(255,255,255,.16),transparent);transform:translateX(-140%) skewX(-16deg);animation:sheen 6.5s ease-in-out infinite;pointer-events:none}
.ov-hero>*{position:relative;z-index:1}
.runline{position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,var(--brand1),var(--brand2),var(--brand3));overflow:hidden;z-index:2}
.runline::after{content:"";position:absolute;top:0;left:0;height:100%;width:26%;background:linear-gradient(90deg,transparent,rgba(255,255,255,.95),transparent);animation:runline 2.8s linear infinite}
@keyframes runline{0%{transform:translateX(-130%)}100%{transform:translateX(460%)}}
.hexdeco{position:absolute;right:-10px;top:-14px;opacity:.14;z-index:0;pointer-events:none}
.ov-illo{position:absolute;right:18px;top:12px;transform:none;transform-origin:top right;z-index:1;pointer-events:none;opacity:1;filter:drop-shadow(0 6px 14px rgba(0,0,0,.25))}
.ov-hero{padding-right:340px}
.ov-illo .flo1{animation:floaty 4.2s ease-in-out infinite}
.ov-illo .flo2{animation:floaty 5.4s ease-in-out infinite}
.ov-hero{min-height:158px}
@keyframes orbSpin{to{transform:rotate(360deg)}}
@keyframes orbSpinRev{to{transform:rotate(-360deg)}}
@keyframes radarSweep{to{transform:rotate(360deg)}}
@keyframes pulseDot{0%,100%{opacity:.16;transform:scale(1)}50%{opacity:.42;transform:scale(1.4)}}
.dsc-orbA{transform-origin:150px 108px;animation:orbSpin 48s linear infinite}
.dsc-orbB{transform-origin:150px 108px;animation:orbSpinRev 32s linear infinite}
.dsc-radar{transform-origin:150px 108px;animation:radarSweep 4.5s linear infinite}
.dsc-halo{transform-box:fill-box;transform-origin:center;animation:pulseDot 2.9s ease-in-out infinite}
@media(max-width:1000px){.ov-illo{transform:scale(.74)}.ov-hero{padding-right:260px}}
@media(max-width:760px){.ov-illo{transform:scale(.6);right:10px;top:10px}.ov-hero{padding-right:205px}.hexdeco{display:none}}
@media(max-width:560px){.ov-illo{transform:scale(.48);right:6px;top:8px}.ov-hero{padding-right:160px}}
.ov-grid{display:grid;grid-template-columns:1fr;gap:14px}
.ov-card{position:relative;border:1px solid var(--line);border-radius:24px;padding:18px 20px;background:radial-gradient(rgba(22,35,59,.09) 1.2px,transparent 1.6px) 0 0/16px 16px,linear-gradient(135deg,#ffffff 0%,#eaf7f2 55%,#e2f2f8 100%);box-shadow:0 12px 34px -16px rgba(22,35,59,.3),0 1px 3px rgba(22,35,59,.06);transition:transform .18s var(--ease),box-shadow .2s var(--ease),border-color .2s var(--ease);overflow:hidden;animation:riseIn .45s var(--ease) both}
.ov-card.ok{border-color:rgba(31,157,107,.45);background:radial-gradient(rgba(31,157,107,.11) 1.2px,transparent 1.6px) 0 0/16px 16px,linear-gradient(135deg,#ffffff 0%,#e3f6ec 55%,#d6f1e3 100%)}
.ov-card.warn{border-color:rgba(192,125,16,.45);background:radial-gradient(rgba(192,125,16,.11) 1.2px,transparent 1.6px) 0 0/16px 16px,linear-gradient(135deg,#ffffff 0%,#fdf2d9 55%,#fbeac4 100%)}
.ov-card.miss{border-color:rgba(207,61,61,.45);background:radial-gradient(rgba(207,61,61,.11) 1.2px,transparent 1.6px) 0 0/16px 16px,linear-gradient(135deg,#ffffff 0%,#fde7e7 55%,#fbdcdc 100%)}
.ov-card.info{border-color:rgba(44,201,176,.45);background:radial-gradient(rgba(44,201,176,.12) 1.2px,transparent 1.6px) 0 0/16px 16px,linear-gradient(135deg,#ffffff 0%,#e4f6f1 55%,#dcf0f7 100%)}
.ov-card::before{content:"";position:absolute;inset:0;pointer-events:none}
.ov-card.ok::before{background:radial-gradient(120% 140% at 100% 0%,rgba(31,157,107,.26),transparent 52%)}
.ov-card.warn::before{background:radial-gradient(120% 140% at 100% 0%,rgba(230,169,58,.3),transparent 52%)}
.ov-card.miss::before{background:radial-gradient(120% 140% at 100% 0%,rgba(207,61,61,.26),transparent 54%)}
.ov-card.info::before{background:radial-gradient(120% 140% at 100% 0%,rgba(56,200,232,.28),transparent 52%),radial-gradient(90% 120% at 0% 100%,rgba(62,207,124,.18),transparent 55%)}
.ov-hex{position:absolute;right:-22px;bottom:-30px;width:150px;height:150px;opacity:.13;pointer-events:none;z-index:0}
.ov-card::after{content:"";position:absolute;top:0;left:0;height:100%;width:38%;background:linear-gradient(100deg,transparent,rgba(255,255,255,.5),transparent);transform:translateX(-160%) skewX(-16deg);pointer-events:none}
.ov-topline{position:absolute;top:0;left:0;right:0;height:3px;z-index:2;background-size:28% 100%,100% 100%;background-repeat:no-repeat;background-position:-60% 0,0 0;animation:streak 3.4s linear infinite}
.ov-card:hover{transform:translateY(-4px);box-shadow:0 22px 50px -18px rgba(22,35,59,.34),0 2px 8px rgba(22,35,59,.06)}
.ov-card:hover::after{animation:sheen 1.15s ease-out}
.ov-card>*:not(.ov-hex):not(.ov-topline):not(.starribbon){position:relative;z-index:1}
.ov-card>.ov-hex{position:absolute;z-index:0}
.ov-card>.ov-topline{position:absolute;z-index:2}
.ov-card>.starribbon{position:absolute;z-index:3}
.ov-ring-glow{animation:ringGlow 3.2s ease-in-out infinite}
`;

/* ============================ root app ================================= */
/* --------------------------- brand mark -------------------------------- */
const PURVIEW_MARK = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAFIAAABeCAYAAACw7kPgAAAdX0lEQVR4nO19eXxURdb2c7o73QlJQETZTNLZw44jILKI7IrigoqioqAg4DisQgBBGNkji4ILCi8IIgqjwyIgyAAuCDgKDAx7QnacGXFEWZJ0J+k+3x99l6q6t1kUhO/9vYdf00nf6rqnzn3qqeecut0B/s/+95t3dOaq+lvn8tX242LMdbUdsLPEMaM/jb4/oZuzbhUAZLzuHZ45mUsrNxa9M2fn1fPO3hxX2wHV0t+fzlUHZXRz1qkCMEIPzSIyqo2vNubmHWlLXrnmUHrNBdKVEAMQhR7QH5oFGQQHIu+oi0YHl3By1svXTECvuUByZdBAImsPw4hCAGWAqkSgSs9U1N+xkL1DM0deHW9Nu6qBTM4axwkDhjaTXgyaSCQKPUzTjmmIJRBcdWJQdegtM9NXvX5V0XlVApk85UVutO8tjumVBMcNkXPV43pEmElCZOBHn8mbrB1HqI27aU00yl3JKfNmXJWA/q6BjOs72NN4/1sc+1QynNd7QqcP4gwAJI7KXNL42CJ2JcSAdESCQIKL+RMm0bn/Obwj8FM5DMQK/wBC1D0paPDdMk4c99Lu33Nsv2sg3Q2q+eg6t8l9KnbcTiDCaSDRrkn+tKltD7d4mko/yUPQH5DbaT9Q9SjEPH1zs4zNizhh4KjWV35kv/fUrmANOY7Qc4jrPKGDJieaSFRWbcFyh42jQ/WfIP+ufwvtTCQDBFdydVQd1XpH2gdvXvHpfsUCmTQ+84B3xAuvSC+yiRxm0hHkBwAmeCxItIOkYjm9h9HprL83r8j9GawFkKXzAO4WN6FhzlpOeXP2FQvoZQ+kd8iIwQ2/nMNVB6Q1ctWNHiUeI6eGNEknhix40tdXRKKA2Aues3jB7D3Hug6gksX/3BE85TNQLaITIER2SUP9vR9z0uSplz2glzWQGaum83UjM+Y5k2JCL5Dpb/0Ns7nKffFgAXLiily8ZK7/zIJD24JllbKGFNqkLsjijLVvcHz/Eal25y+YNqnt4VufoLL12Qj6A5Z+GADFeBD1SBNkfPUhe4ePm3y5xn7ZAun947AHIpvXgMF/Fn4jUIwbIJP7iAhg9uktCqbP6FR+4JSkE6U+AoyIRjVRbextOekfhdeNucPG0OGGD1P5t99L/ZCAdFftqoge1GZ8+tolHNd3hOe3jv/yIdLliNd/tF2RRf1ntLNZTGx0IqT2AJjg/kNtNDz+MacufjVsQHN6/4nOzNzevCLvlOFXyDcy+nLVq4XYMV18KW//NkH/qwKZNCGzMHH4C9K0YIZfQprKbWRdVTV0RNq7RQLXaV04NTQTARRCfmQ7Lxoe/IiT52bZBqJ4waw92Xf1pXNL9+7hX/xapmRd5T3t01Hv27WcOG7Kr9KflxzIBn97hav2T0pw1PaMF18ngocFTmJlWMHyoHBMXFnNqQ1AQ4t9P0GfyZ8S/3mciLo7HfX3rOTEKVNsA1o4dWLzIy0fptINR8AVMn/qk4BioxDVu2Wz9C0fccKwl8ZdSlwuOpCpr4/nJsffZHe9qiFEsImURvve5KrPps4zMwyRJ0NWsrYwsjL7zHl1YuO8D9jTsramJ4XVW7O8oWPJ/82/zL6NXDz0cFSNRPSjjVFv+3KOf37Mo3bjyB8+ko40upfKdxcJrCnzuvOm6oge2H5K2voVHD9wzEUJ+osOZPQDN4HcDoNjpHcyw1HVHcp7AVknanZiyTz/2Q9zIy06UerGzJ11JKrwyuk9lE5n7WpUkXNKWPm1IGi+OWrFourQNitS338rLO8d7z2Qzs7+vFFl/iljlqjmTK6J6KFdd6QsW3xB/rz4qV3BIV4inVfEt5qrrKQTFd9OLJnrt+hEsReb3NkusyleOOvQsW79qWTRvo3BUz5InCes0p5bvWiwfw0nZdkXgosXZh3K7vYYlS77dgefEfsxfSQ44G6WjIy9Wzhp+pywAbUEMv6ZITfW3zDd8gYOwuQVACxgxcJp4kPtB0I/LDdROetCmU3BjEl3H2n5GJWuOQwurQg1V32JjEDU/Y1Qb8fH7M2c+KFdP4XTxrc92vJe8m08AK4IgJllXxggTwQ89zVD2rZ1HP+nl55S+5ACmf7uRK4xKeOku16MzenMq0UgICgc8rGSiYRHk1pP1FNE+Rz2fcQPeEGuXWqWP2osHW76EPm3Fxp+mNwZmiGOG2IQ/UyrXmnr3+e4Z0beaNvPiOF0tElXqthdYM4sxRdnresQ/dydS1Pef1+6xA4ASBz7wtam2fM4qnNNSWMBQPqyyZyxYiqT22m8xizHqGR10dKgT6hs2ziZsXI2Z6yZy6pONIsWQMnKo9oJSOtHDqT7lrq76//9A/a+OOFru0Dk9htMp6dti6w4+IM0BjGzcaXeiNhRXU+mLF0QXtA/1Y/OvbqxUSD/v9JgQrwf4nDXzUnI+MdX7H1xxtdGIKv2SehIVVyG/hOr0kQORLa+EZLuUnRiwcxX+p5dlLNC1YmSRUbA3fQG7X2iTjQt/8WJVL7nB0knqtmR4/oqiHm6WZuMje9yfP+RllTxxJI5/pwH+1HJO7tmBn8skfwxVmcieFomo96+TZw00573ihdmHcq552EqW/lNEZdWSOM2ODzChajH72hjBJKiXTJviVdB6DycRgQA9gf2qDpRbqDqP7JwJACNoyD5IzmgPblSaiA2s11O6nv2JbLCOZMzj7btSWVr/wkuDwj8SaYvbhci72mK9O2fsHf0pFW2/Uwa7T3WojOVf3tc5l9xTEYgBW6zoCAIiEjTdaLEkaEmHvt6on5cQIai/+RICj6oezaKjwSC5zYvGhxcy8mzZ9oGND8zk440uYfKvy0UzifPLGeNWFTp065H6poPOa7vSNu8O+/pp6l04ZZROFMmcKc5Rgcgr7bWPJkFjjF1ojp+ioroeL7KNgeCtjrRYgG2veIAtIyEhFxdOx7hQmT3hsjY+TF7x7681a7b3KcG0LnXtrYJnDhl9qf460qrg5hR9/qSF9rzZ/HcP8+qKPzR4HBt5AB0RBq6zmHhP9+uU90C//FZdaIQyQbb3+KoDnU7WuqJQhv/F98PYX/QRiealrZ8HrtSrrepJ4Ysb/Aoqiz8RUC2Q0K2s0Y0ovu27Ji+8X2OH5BpWeGL35mxM7vLI1S24ruTwbN+YzySP0Rwt66Penu2ceKM1ywBJf1/kjMze0QKzhfNn7Pp7LL8PhadKHYeHQFXWjVJ/6ntCue98nrp2uN2OtGQP+5bb4KjVoyNP6Yd6/IUlbz/j6LAT2UW3ar/7ky6ATEjOu1OWfKOLbIKXh5b69itd5P/y6Oy9hR8Z08EPN2bI237Z5wweprBn5Z2IkfK+k/hNgBcUrnGohMlF5XcOUxmA3/gtEUjEkxOUrmTdPQrgZg0wXu0VU8q++QQ2B+QdKmIMnfLZNT/5yZOmmNfast77jk62qA9VewrgsS9Qq7vrB6LqCfb90hdv4bjnx3bUOVoncNDgdT0kRk55YzafoqkE9U1Qv9PQYitiTpRFOTSjNDPZyfqQ5Y/KpOONLmP/NvzDL+k1RkI8eddjZGx4xP2vjjFXn8+8RSVvPFZt8CJUxCrT+a4CE5vLVQZcv9BV1qcJesBjMxGRYmN8zb1RGP8AUBGrLVyI3Sk6ETBKvVVTH2c33L7P0+H691FlYd+MDhcRbbj+hhU6d2mTdqGv3Bc/9EW/Vk8f+qmnDsfIN+KHQf55xJhjMLsAIEi3ZLq0McgpYhGpC3ShiJlnUlgAW6la/L2BE77rTpR6KL+F4vY3aJuNVVPioEqeW/fNvYLOhLy5EiZP5dTl78dDufIeagPnVu4fV7gvyUWnjX5syZihnfLSVnyrj1/ThndOPv2O8m3/jtwWYU5A5V+VA6XEBluNdXbSDpRmHIFU6c1L12Tf9qqE01zXB+FiIwbICFaaVMwfXKnigM/yBwkIh+Au7kXDQ5u4KRX7TOSotmThma3fYB8Gw8qmYhD0rIRLVKQse9zTpw1zz6gY4ZSdos7qGLXERtONFVOCKnCqi3pxDBboFI2otSNAj+VZZ5fJ+oIlfedLefwBaw60TioPbkciOrWABk7V7F33CTbrYH8EcPoSP1OVL6/2KomdHO5EHnXLUjfvokTxkz71K6fvAHP0LHGLaniYKHAiaTzu+GTlNmIK54aBBMh+uqsRJLJZ19PNPuQEBsO+cxCOwfkvF/maOf1MYh+slWz9HXLOb7vC7YZSW6vPlT6zrZRwZNnYc4CBwBTMzuqx6JK7/bdUj9ZzXH9xsTb9ZP3WE8qe3tDD/gDRjzC6siwubS+HwPhytrAyVJPFI8Fg7YcYzGnI6wvXC7f62PwXmpNxIzu5ktZZK8bi+ZOnpXd/l4qW/0d+Jxfex/L5wHgTKqF6GH3FyUvXmbbT/FbL68JFP0gqRe9oUN/klZdSxTgl/dYyJprA7C7UrqVfZID+AJWnajJn4Q/Zj6Qsfk9diXXUFZdE/kV+76vxufKoe6x6A9361TU272JE6fYV8QLx4+kY7d2oYrth+U9IaXuGNEiHel7d3LiTJutiqCpLEQO13QkIOpE61YqPAansfFaeGPrxcifMJHKtuVbdaJejyTyOL3V4agZC6MmqvRR/O6sM0eb30f+HXlgG93LTEAVNyIfbIG0bWs5fvD4AXbu5T03kI42bEOVh4oNnQjBL2YALhfcdzZH2pefc3zmdKOyLobgApmNPXepOlFdlfVWok5U+wn+VFZ0Xo2o5M5qTq9bbr9BdG7mZzUrs09C0ngCx7tqX4foQV3eSVm5PKxcyn20F5W+valP8ORpqHeA6H44qsciqnfXXilr13Ncv7HxIv+LHG4gMlw9Ujc7XWZtZB63J0r2SzpSeav63vNViYoXz/ox5/4nqGTR9gXBH89asg39fRGN4pFxcDsnzZtvz3tvTHovp9Nd5PvL9qLguTI5BsKzM7Euoob0LHJ4a9tyeGixCcpXgkhekYnIY6knWtBEkXJGQ2Dr3RYesfJNIMBB1QCAjDs1zF1IVUfaWdGsPw881u5+8q3ZCy6rsOVOIoKnUxOk7/wbe8dl2cqlwkkjvTmt2lP55/uBSn0PyiEjFQTyeCDqSR2LIfmjxY01buKgHSS149B1otKGyKPqRFK7CeKMiERmAEGcBoCit7LWBP5zxnydBV2rWcILE+Ymv2O/V13w4kg61rwr+bYfNV0W9liYAapaBZGP3d4sdf0aju8/1v6OtqHPUvYtLanycJFlVpmoFzhc4khpH9eOl0xeNP4pqLXXiYo5qKrczmEuNgCO3dGLAkW/SIgV98/J6azpaZeBevs2c9Kr9h+tyx84kI42vIMqD39v3WPRnp3e2qgy7IGccDIHAHJ7PUhlSzZN4dMlgOCvnNnoPK7qSPkCSKZqrvPqyHDcRohkpR2CfEZsUlHwk0UnyicB4HbBc2djpP99I3vHTbWdqsd7PkYlb28eWPnvX6RMSvzZ1Twdad/tYO80+zSx+NVxL+W0u538a7+GyJ8Sf4uIlDMWG0Qy+y13cdktJJa9GCuy9StptBPrkQAoAFh1otjARKojNgpRT7RtlvrJR/aBeH3SguOd7yH/ql0+PmveSSGi0+Fxw9O9FdK++JwTRkx9xa6fwgmD6XibVlS++RuoHK77F0IkzKtkB0t2UDU55zUhLQzQY6knqsbwha6ixDF+pY1RPbIzNcMCA66Umqh36CtOes1+R7Fg4sio7FYdqXzLfhgcLs1AAl1fFZF97hqV8te1Yad7YeZACuQUm6pC4PAQAQUVnWjRiCQjxEbfEcGjrpY2q3akRSc6qKrUxqEgVjqPA3Z7LDrHe7o0RcaOzex9KSvbLhD5wweFChB7c7VelXGB4ExNQNrePZz0xlL7gFaaubbI4RpHkjL/rWgQERA2T1aP2y3+ik60NRsOkk4hvtdop2Uk1aIR+UjbtNQtn3L8H8c/YNd9ft/HqXTOXyMDBT/I/uinczjgatsUqTt2sXfC3ELpzUS2HG7WI8Ps3Ak94LyVbya/rrtMtNgNQ+EYdWoL6LC981dBkNzO5GBn7eqo8lz31SnLP7TnzyVZ/tz7ulPZe5vnBf97GibCtOyICBQdDfeDHRNSPv2c4wa81AwAKCJC4nAp1yYHyTrRUpBg44rpHGNtwdKVBRCmsKHoRBuONNuRMjvYcrxifyGCp0sNn0Izx1wMXU2Skb77a/ZOfs0+oLNfHHq8U0fyr98JLvUL2ZGpEx11bkDUc4/uTlyyihETLXG4zJGQdaItCsSMxC4H1jSirBOVANqsmiToSMMVi67Vj4UyLLGeWLE3f0F2665UvvWA1LeUkXjciHygDVI2fMpx/cba1hsLxz1POa1uo4qv9kFEtliDiGhaD46aNaRcXF61Ff1nQaS22p5XIwIW7rAkP/oxqR375T6Ue8jlg0aubvYfapE/dBAda9Sayv+RJ2tG4eGMuxFVhjxUlDQ/zEICoGBIX8q5uQlVHimQagLquNR1wChaiBEmh5XcLPrPGkU/CVeSjAqQ2Al5DA4y0a1UtoXV0HIuc0bY1SsBIO/JJ6hs/voewX//LK3upKOUCBGtGyPt6x2cMPoV2xunACDvsXvJv2zDzOAvZxGu/mmqBsELqRJtw21SNcR+FRGulJaz27Ux2pEeG8sWgaoTDSN41H1nOyt+a8qanK53ke+v208Gz4WEuLnHEvKLYqvA81iXHinrPuP4oZNtv4GgeM6YzNwOraj8i29D32ygOGj0CWPPRo64uioT+AIZizZIpZ5IltGaSNPbsfZ5baGJhCSJaBl+NXe212EhK3z5hVo5rW+niu0HJN/FXN+ZUBuRT98/M+XTrRw/Yrp9ZjO8Lx1v0YiCucUw+NdQOaKOhIKAMNoNF0LlhTSk0I+BKsvUVvhV9UXkzvPMDtHyn3+aspvcQhUH88NynaNOTXie7D4qee3f+KZ+4+w3wHp2peC/ToavR8L4KNv5PiOtXwm7XUaAgnTG6ENb7Vitaxr/i1ynTm015zfPxQ6qZqoKKz9eyPKfeIh8i9YN4R9Pyxwr6FpnQl1EPf94UdLiv9pCnUvKJA6XK+QgmPcdhrvCpq4SFkvzBMIL+lVWpzYrP4Rqn8rUNjaXdC0n+6PqxEu14nkTXj/epT351nzp4zK/NAvN3JngvLkBUvceYe+sd+UzBYLSDFYyGwhz3i6QMkpstSaRRzoOa66taryQjlQQ6TD3ze0yGzud+GusaOKQqOOtWlDl3w8YFCFlNRpiIzq2RvKW7zh+WNY43T+LGoDBkYp2C9ic2UZPKcf9alWFgpYurPpORSSzLQdpHfjU2iLUzzJeouUPepxybmlMgbwTYXUiVa8G95M9piR98BlTTIzkl+5eaNVmWSdadCSz37xiOp9YTcpVbREJmKuuxjFq9YdkRMhnkveEws+gS7e8h+4k//vrZvKp01B1ou6vIyMFjjq15OMyItXVVA2k9r0Tok60GYDOMeHqicYxCKiyaWR3xU2T64mX04rnjM7M7XwblX+2HewvN3SiHSeGnDM53LzT4nw6Ua81SnsWlvF5JKSF2fdW64nWflSdqLZQM5vLg0jRisY+S7mtGlNg7yHon7KQ/dH4Wd2zAUykhUlH/FJmw/IqLbYR21l2ERl+lWMtHAn5Uw1hdaQxMy43Lk0reLYH+ea/lxAo+pec+enntq7aQhZBDlCU29KpqrvCsWRYVIdrowpyYxtD1LWGE5ESd4b14/LZiUWTi/N7tKOKj9dt5DPCJ8mUfXc5kIJOTFkob3eqn+2zrtok3Qse1lg/n64TleqPZWaQeNCnZ1Umf105RIpWNGPY3eXLP04w9LbC0y7TWd3hUIwj70hBxmfLuHT1wY6oCBxXj1sRxxBQZrQVjRg+kE7QpDWz8LHZB9v0oQdRf128gFfQvHOWs6tdm5AHul9kRsMFAIH/nIOzVqwl83AmXo/Y4e22Veadsq6wNjOKOfRyOIzoVXQZzRauDV0SO9RDrs6HxsO/+atmzmfx497YHdGlfTOKjrHN/4MlJQC0qX1u8e5qge/PmHNeqfe5km+QOIHgsEaL4bPcs62YWXU5z6rLQjvYcKBYXbpCqzYAxA16uUPy+m/Y3eOeZhQdI3Oi5n+woAj5HeoRoAXyxLuvnjnaoTeVrNiPYEkF1PsTpT0W/XfLPgL5z/dZRK0fv7WeaKM3Bf6RZwFFqp9FvBKB9C7ZwJ7+T26jWrWgn0/Kes6cQ/nKVQsKHmlrnFz61ueCCeMJE4CUha+zp12qwVc6J5o5NOBucpN8dgeqyQixooUIobvajAUDsIRcrFUyFESK2lNvd/kWmxAPtjJ918ev+0tA5bYvUTT6ScvVs61D5T47mM5mba1ZcUyrvUG+KmDAmVILGbs3sPelqdkAwNqKql490VjJx8PFQK60K8fO0/+vtfixc7cmf/FPdrW7zTi/WhOoPHAIeS0SyC6IwHkKesWLZ/2Yc18fKn131+rgz6Ww3htEcERHosrjrdLSt67miAZxS0HycWtyI1aObHQiYK7oejuppilXaH5rZhM3YGKzpHU72P1Q944UE2P6LehE/s8P8C9a2qbwmW7nPdEFv9C9MOvlB5EFJL32Kns61Ae5XYbS0b/bzFHnOrjrVLdov6AySDNnDe8Ta+8X90MMI/aYK6eue3+d/Elc/Ak7mzSAdDG08TAAnCtB5dYvc4omP5d+Mf1d9Dfj5w8bTgCQumIJRzRNAGCurtB+C1GJoBNV+oOgA8kmUICp0SSu1YzJTxRaCEnjrEuVP95ZSzmiQ0gP6hxs6FKNEyt3fouiIT0vCeqXVqsHcLxXXzr35pY+gaJTsHAiIHGLq0E84geONb4Sy2grtLNOTSHDUrUrrPx1sYhMGPva1uQv9rGrfRtLdUnvM5Cdh9zmXrrUIAK/8tv6it+Y9l72nQ9T2Qc79/AvZVB1na4TnSl1ED343h2pH/0lFLIgn5HawWbVFRBpufOX4LHcO3QBRMY9O6Fh8tqv2P1w944UG2NT5SLg1M/wL185vuDxjr+acH/TH70onDquOaYCSbPmsbtzY8Dl0qaKaQzAWS8eadu3cPCHX6BmJ2BYKtx6smldsWU60O7lDovIxP9Zza4/NApb/+TySlR+/jWKxj3zm8XoZfnrIfkjhxAAJC9azBG3pkH/VJW+cAChz6s4rouFmGMzw/hUg2FG/mqtaRorOeu/2Vd/vFmLOeKOlkCEW3+X7A8Dgf2HUND/3sum5n/V1A5nef2eodLX1iVUZP9bqieq3CnqRItZ2omLjfY5Hb1rRYwmjH11a8q23ezq3BYcEWG79xM8+RN8C5c1v5xBBK7A37MpXvRKMRaBEl6YPNfTrcUQR63rBE4EQMJqSQS6LrYPgL5GBwzoebr6lV4hEzIs7ZW4/uNT3fd0ynEk1JVUA2voZgZQ5kPF5i9ziib/8aLkzKXaZUWkaEWzXxqa0/ku8q3dBS71G8hU64me7m2RvHI1xw0Y3wwQc2xz9TaM4ZMyLACu9reNinyud44joa72HpI5kYHKb/ch9/aGdKWCCPwOf2GpcPxwwngg+Z2l7GrVUFunZZ3oqpcIV72k3Um3NoCjRnXoqGONS03T94RCPzMDzrg6RhtiAfkgBE/8C+XrNiecWDS5+EqP83f7U1V5A/tQXN9Mj+e+tj5nSqjgwcbyrDnTvGHoB4EHpYntdNRU904kAaVrw3OlqNzy1cGiKX9qfGVGY7XLX4O6CEvInP6h++7WvRzVq4ZcMG4jVjIajQZy/tCQACDpL5vYmZooHxc5EUDlzj0oHPzo7z6uqxJI3bwz5rO7UwsgIgJi4EJm7p1XfHMAEQ1SgNgYqR3rFAAgmHcC+T07XLXxXNVA6pa09GOOaJoh6Twp89CQx9LrCC0sZ86hfNPnq4uzRjx4FVw37JoIJADEP//nRyPu67DCWbOGgUTzdkPAgsRAABU79qBweO9rYgzXhBOieafM54jOrUAej1LUMJEZOF6E/Ec6X1O+X1POiJb47ip2Na0PCYlnS1C+6YuNxTNG3H2V3fv/y+IGTeyQvGknp369nxOmhP9y4f+z/0X2/wAKH5U6/NRM8gAAAABJRU5ErkJggg==";
function HexLogo({ size = 20 }) {
  // The real Purview Healthcare mark (keyed out from the official logo), not a redraw.
  return <img className="logo" src={PURVIEW_MARK} width={size} height={Math.round((size * 94) / 82)} alt="" aria-hidden="true" style={{ display: "block", objectFit: "contain" }} />;
}
function Wordmark() { return (<span className="wm"><span className="wm-grad">Drishti</span></span>); }
function Brand({ size = 20 }) { return (<div className="brand"><HexLogo size={size} /><span className="brandtext"><Wordmark /><span className="bytag">by Purview Healthcare</span></span></div>); }

/* ===================== floating feedback widget ======================= */
// Pure logic (mirrors scratchpad feedback-logic.mjs — kept in sync and unit-tested there).
const FEEDBACK_KINDS = [
  { id: "bug", label: "Something is wrong" },
  { id: "idea", label: "I have an idea" },
  { id: "praise", label: "This helped me" },
];
const FEEDBACK_ORG = "Purview Healthcare";
const FEEDBACK_RECIPIENTS = ["purviewautomate@gmail.com"]; // one place to manage alert recipients
const SCREEN_NAMES = { overview: "Overview", dashboard: "Client detail", clients: "Clients", formulas: "Formulas", guide: "User guide", questions: "Q&A and Feedback", users: "Users", audit: "Audit log" };
const friendlyScreen = (view) => SCREEN_NAMES[view] || "Drishti";
const maskLongDigits = (s) => String(s == null ? "" : s).replace(/\d{9,}/g, (m) => "•".repeat(m.length));
const trimToCap = (s, cap = 1000) => String(s == null ? "" : s).slice(0, cap);
const storeStr = (v) => String(v == null ? "" : v);
const movedEnough = (dx, dy, threshold = 6) => Math.hypot(dx, dy) >= threshold;
const clampTab = (x, y, tabW, tabH, vw, vh) => ({ x: Math.max(0, Math.min(x, vw - tabW)), y: Math.max(0, Math.min(y, vh - tabH)) });
const initialTabPos = (vw, vh, tabW, tabH) => ({ x: vw - tabW, y: Math.round((vh - tabH) / 2) });
function placePanel(tab, panelW, panelH, vw, vh, gap = 10) {
  const tabCenter = tab.x + tab.tabW / 2;
  let side = tabCenter > vw / 2 ? "left" : "right";
  let left = side === "left" ? tab.x - panelW - gap : tab.x + tab.tabW + gap;
  if (side === "right" && left + panelW > vw - 8) { side = "left"; left = tab.x - panelW - gap; }
  if (side === "left" && left < 8) { side = "right"; left = tab.x + tab.tabW + gap; }
  left = Math.max(8, Math.min(left, vw - panelW - 8));
  let top = tab.y + tab.tabH / 2 - panelH / 2; top = Math.max(8, Math.min(top, vh - panelH - 8));
  return { left, top, side };
}
function buildFeedbackEntry({ kind, message, name, org, email, screen, now }) {
  const t = now || Date.now();
  return { id: t.toString(36) + Math.random().toString(36).slice(2, 6), when: new Date(t).toISOString(), kind, message: trimToCap(maskLongDigits(message).trim(), 1000), where: screen, name: storeStr(name), org: storeStr(org || FEEDBACK_ORG), email: storeStr(email) };
}
function buildFeedbackEmails(entry, recipients) {
  return (recipients || []).map((to) => ({ to, subject: `[Drishti feedback] ${entry.kind}: ${entry.where}`, body: `${entry.message}\n\n— ${entry.name || "Someone"} (${entry.org})\nScreen: ${entry.where}\nWhen: ${entry.when}` }));
}
function makeRateLimiter({ minIntervalMs = 1500, windowMs = 60000, maxInWindow = 12 } = {}) {
  let times = [];
  return function allow(now) { now = now || Date.now(); times = times.filter((t) => now - t < windowMs); if (times.length && now - times[times.length - 1] < minIntervalMs) return false; if (times.length >= maxInWindow) return false; times.push(now); return true; };
}
const fbAllow = makeRateLimiter();

function FeedbackWidget({ me, view }) {
  const TAB_W = 34, TAB_H = 132, PANEL_W = 320, PANEL_H = 380;
  const [pos, setPos] = useState(() => (typeof window !== "undefined" ? initialTabPos(window.innerWidth, window.innerHeight, TAB_W, TAB_H) : { x: 0, y: 0 }));
  const [open, setOpen] = useState(false);
  const [kind, setKind] = useState(FEEDBACK_KINDS[0].id);
  const [msg, setMsg] = useState("");
  const [sent, setSent] = useState(false);
  const drag = useRef(null);
  useEffect(() => {
    const onResize = () => setPos((p) => clampTab(p.x, p.y, TAB_W, TAB_H, window.innerWidth, window.innerHeight));
    window.addEventListener("resize", onResize); return () => window.removeEventListener("resize", onResize);
  }, []);
  const onDown = (e) => { drag.current = { sx: e.clientX, sy: e.clientY, ox: pos.x, oy: pos.y, moved: false }; try { e.currentTarget.setPointerCapture(e.pointerId); } catch {} };
  const onMove = (e) => {
    const d = drag.current; if (!d) return;
    const dx = e.clientX - d.sx, dy = e.clientY - d.sy;
    if (movedEnough(dx, dy)) d.moved = true;
    setPos(clampTab(d.ox + dx, d.oy + dy, TAB_W, TAB_H, window.innerWidth, window.innerHeight));
  };
  const onUp = (e) => { const d = drag.current; drag.current = null; try { e.currentTarget.releasePointerCapture(e.pointerId); } catch {} if (d && !d.moved) setOpen(true); };
  const closePanel = () => { setOpen(false); setKind(FEEDBACK_KINDS[0].id); setMsg(""); setSent(false); };
  const send = async () => {
    const text = msg.trim(); if (text.length < 3) return;
    try {
      if (!fbAllow(Date.now())) { setSent(true); return; }
      const entry = buildFeedbackEntry({ kind, message: text, name: me?.username, org: FEEDBACK_ORG, email: me?.email, screen: friendlyScreen(view) });
      if (REMOTE) await api("feedback.add", { entry });
      else { const list = (await sGet(K.feedback)) || []; await sSet(K.feedback, [entry, ...list].slice(0, 1000)); }
      // buildFeedbackEmails(entry, FEEDBACK_RECIPIENTS) is ready for a backend; offline build stores + exports instead.
      Promise.resolve(appendAudit({ user: me?.username, role: me?.role, action: "sent feedback (" + kind + ")" })).catch(() => {});
    } catch { /* never let feedback break the page */ }
    setSent(true);
  };
  const left = msg.length, remaining = 1000 - left;
  const panel = placePanel({ x: pos.x, y: pos.y, tabW: TAB_W, tabH: TAB_H }, PANEL_W, PANEL_H, typeof window !== "undefined" ? window.innerWidth : 1280, typeof window !== "undefined" ? window.innerHeight : 800);
  return (<>
    {!open && <button className="fb-tab" style={{ left: pos.x, top: pos.y }} onPointerDown={onDown} onPointerMove={onMove} onPointerUp={onUp} aria-label="Give feedback"><span className="fb-lab">Feedback</span></button>}
    {open && <div className="fb-panel" style={{ left: panel.left, top: panel.top }} role="dialog" aria-label="Feedback">
      <button className="fb-x" onClick={closePanel} aria-label="Close">×</button>
      {!sent ? <>
        <div className="fb-head"><div className="t">Share feedback</div><div className="s">On {friendlyScreen(view)}</div></div>
        <div className="fb-body">
          <div className="fb-kinds">{FEEDBACK_KINDS.map((k) => <button key={k.id} className={"fb-k" + (kind === k.id ? " on" : "")} onClick={() => setKind(k.id)}>{k.label}</button>)}</div>
          <textarea className="fb-ta" rows={4} maxLength={1000} value={msg} onChange={(e) => setMsg(e.target.value)} placeholder="What's on your mind?" />
          <div className="fb-meta">Your name and the screen you're on are sent with this. Please don't include any patient details.</div>
          <div className="fb-foot"><span className="fb-count">{remaining} left</span><button className="btn primary sm" disabled={msg.trim().length < 3} onClick={send}>Send</button></div>
        </div>
      </> : <div className="fb-thanks"><div className="ic">✓</div><div style={{ fontSize: 15, fontWeight: 600 }}>Thank you</div><div style={{ fontSize: 12.5, color: "var(--ink-3)", margin: "6px 0 16px" }}>Your note reached the team.</div><button className="btn primary sm" onClick={closePanel}>Close</button></div>}
    </div>}
  </>);
}

function timeAgo(iso) {
  const s = Math.max(0, (Date.now() - new Date(iso).getTime()) / 1000);
  if (s < 60) return "just now"; const m = s / 60; if (m < 60) return Math.floor(m) + "m ago";
  const h = m / 60; if (h < 24) return Math.floor(h) + "h ago"; const d = h / 24; if (d < 30) return Math.floor(d) + "d ago";
  return new Date(iso).toLocaleDateString();
}
function FeedbackAdmin() {
  const [list, setList] = useState(null);
  const [tab, setTab] = useState("all");
  useEffect(() => { (async () => setList((await sGet(K.feedback)) || []))(); }, []);
  const kindMeta = { bug: { label: "Something is wrong", cls: "bug", chip: "miss" }, idea: { label: "I have an idea", cls: "idea", chip: "warn" }, praise: { label: "This helped me", cls: "praise", chip: "ok" } };
  const counts = { all: (list || []).length, bug: 0, idea: 0, praise: 0 };
  (list || []).forEach((e) => { if (counts[e.kind] != null) counts[e.kind]++; });
  const view = (list || []).filter((e) => tab === "all" || e.kind === tab);
  const exportJson = () => { try { const blob = new Blob([JSON.stringify(list || [], null, 2)], { type: "application/json" }); const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = "drishti-feedback.json"; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000); } catch {} };
  return (
    <div className="card" style={{ marginBottom: 18 }}>
      <div className="row" style={{ justifyContent: "space-between", marginBottom: 12 }}>
        <div><div style={{ fontSize: 15, fontWeight: 600 }}>App feedback</div><div style={{ fontSize: 12, color: "var(--ink-3)", marginTop: 2 }}>Notes sent from the floating Feedback tab, newest first.</div></div>
        {(list || []).length > 0 && <button className="btn sm" onClick={exportJson}>Export</button>}
      </div>
      <div className="seg" style={{ marginBottom: 14 }}>
        {[["all", "All"], ["bug", "Something is wrong"], ["idea", "I have an idea"], ["praise", "This helped me"]].map(([k, l]) => <button key={k} className={"segbtn" + (tab === k ? " on" : "")} onClick={() => setTab(k)}>{l} <span style={{ color: "var(--ink-3)" }}>· {counts[k] || 0}</span></button>)}
      </div>
      {list == null ? <div style={{ color: "var(--ink-3)", fontSize: 13, padding: 12 }}>Loading…</div>
        : view.length === 0 ? <div className="empty" style={{ padding: "36px 20px" }}><p>{(list || []).length === 0 ? "No feedback yet. It will show here as soon as someone sends a note from the Feedback tab." : "No feedback of this kind."}</p></div>
        : view.map((e) => { const km = kindMeta[e.kind] || { label: e.kind, cls: "idea", chip: "info" }; return (
          <div key={e.id} className={"fbadm-row " + km.cls}>
            <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start" }}>
              <span className={"chip " + km.chip}>{km.label}</span>
              <span style={{ fontSize: 11, color: "var(--ink-3)", whiteSpace: "nowrap" }}>{timeAgo(e.when)}</span>
            </div>
            <div style={{ fontSize: 13, color: "var(--ink)", margin: "8px 0", whiteSpace: "pre-wrap", lineHeight: 1.5 }}>{e.message}</div>
            <div style={{ display: "flex", justifyContent: "space-between", gap: 10, flexWrap: "wrap", fontSize: 11, color: "var(--ink-3)" }}>
              <span>{e.name || "Someone"}{e.org ? " · " + e.org : ""} · on {e.where}</span>
              {e.email ? <a className="link" href={`mailto:${e.email}?subject=Re:%20your%20Drishti%20feedback`}>Reply</a> : null}
            </div>
          </div>
        ); })}
    </div>
  );
}

export default function App() {
  const [booted, setBooted] = useState(false);
  const [users, setUsers] = useState([]);
  const [me, setMe] = useState(null);
  const [view, setView] = useState("overview");
  const [clients, setClients] = useState([]);
  const [dashClientId, setDashClientId] = useState(null);
  const [dashRange, setDashRange] = useState(null);
  const [problem, setProblem] = useState("");
  const [authNote, setAuthNote] = useState("");
  const [navTick, setNavTick] = useState(0); // clicking a tab again always brings you back to its first screen
  const loadShared = async () => { const [u, c] = await Promise.all([sGet(K.users), sGet(K.clients)]); setUsers(u || []); setClients(c || []); };
  useEffect(() => { (async () => {
    if (REMOTE) {
      if (tokGet()) { try { const r = await api("me"); await loadShared(); setMe(r.user); } catch { tokSet(""); } }
      setBooted(true); return;
    }
    const stored = (await sGet(K.users)) || []; const aligned = alignRoles(stored); const u = aligned.next; if (aligned.changed) await sSet(K.users, u);
    setUsers(u); setClients((await sGet(K.clients)) || []);
    const s = await sGet(K.session, false); if (s && u.find((x) => x.id === s.userId)) setMe(u.find((x) => x.id === s.userId));
    setBooted(true);
  })(); }, []);
  // shared-server housekeeping: an expired pass sends you back to sign in, a failed save shows a banner
  useEffect(() => {
    if (!REMOTE) return;
    const onAuth = (e) => { tokSet(""); setMe(null); setView("overview"); setAuthNote(e.detail || "Please sign in again."); };
    const onProblem = (e) => setProblem(e.detail || "Something went wrong.");
    const onSaved = () => setProblem("");
    const onLeave = (e) => { if (pendingSaves()) { e.preventDefault(); e.returnValue = ""; } };
    window.addEventListener("drishti:auth", onAuth); window.addEventListener("drishti:problem", onProblem); window.addEventListener("drishti:saved", onSaved); window.addEventListener("beforeunload", onLeave);
    return () => { window.removeEventListener("drishti:auth", onAuth); window.removeEventListener("drishti:problem", onProblem); window.removeEventListener("drishti:saved", onSaved); window.removeEventListener("beforeunload", onLeave); };
  }, []);
  const refreshUsers = async () => setUsers((await sGet(K.users)) || []);
  const refreshClients = async () => setClients((await sGet(K.clients)) || []);
  const login = async (u) => {
    if (REMOTE) { setAuthNote(""); setProblem(""); await loadShared(); setMe(u); setView("overview"); appendAudit({ action: "signed in" }); return; }
    setMe(u); await sSet(K.session, { userId: u.id }, false); await appendAudit({ user: u.username, role: u.role, action: "signed in" });
  };
  const logout = async () => {
    if (msSupported()) msx.forget().catch(() => {});
    if (REMOTE) { tokSet(""); setMe(null); setUsers([]); setClients([]); setView("overview"); setProblem(""); return; }
    await sSet(K.session, null, false); setMe(null); setView("overview");
  };
  const setFirstName = async (firstName) => {
    if (REMOTE) {
      if (me.role === "admin") { setMe((m) => ({ ...m, firstName })); return; }
      const rest = String(me.name || "").trim().split(/\s+/).slice(1).join(" ");
      try { const r = await api("users.updateSelf", { name: [firstName, rest].filter(Boolean).join(" ") }); setMe((m) => ({ ...m, ...r.user })); } catch (e) { setProblem(e.message); }
      return;
    }
    const all = (await sGet(K.users)) || [];
    const next = all.map((u) => (u.id === me.id ? { ...u, firstName } : u));
    await sSet(K.users, next); setUsers(next); setMe((m) => ({ ...m, firstName }));
  };

  if (!booted) return (<div className="app"><style>{CSS}</style><div style={{ padding: 60, textAlign: "center", color: "#8a94a2" }}>Loading…</div></div>);
  if (!me) return (<div className="app"><style>{CSS}</style>{REMOTE ? <RemoteAuth note={authNote} onLogin={login} /> : <Auth users={users} onUsers={refreshUsers} onLogin={login} />}</div>);
  const can = (a) => me.role === "admin" || PERMS[me.role]?.includes(a);
  const activeClients = clients.filter((c) => !c.archived);
  const usersTab = REMOTE ? (me.role === "admin" ? "People" : me.role === "manager" ? "My team" : null) : "Users";
  const tabs = [["overview", "Overview"], ["clients", "Clients"], ["formulas", "Formulas"], ["guide", "Guide"], ["questions", "Q&A and Feedback"], can("users") && usersTab && ["users", usersTab], can("audit") && ["audit", "Audit log"]].filter(Boolean);
  return (
    <div className="app"><style>{CSS}</style>
      <nav className="nav">
        <Brand size={20} />
        {tabs.map(([k, l]) => <button key={k} className={"tab" + (view === k || (k === "overview" && view === "dashboard") ? " on" : "")} onClick={() => { setView(k); setNavTick((t) => t + 1); if (REMOTE && (k === "overview" || k === "clients" || k === "users" || k === "questions")) refreshClients(); }}>{l}</button>)}
        <div className="spacer" />
        <div className="who"><span className="who-mail" title={me.username}>{REMOTE && me.name ? me.name : me.username}</span><span className="rolechip" title={me.username}>{roleLabel(me.role)}</span><button className="btn sm ghost" onClick={logout}>Sign out</button></div>
      </nav>
      {problem && <div className="probbar" role="alert"><span>{problem}</span><button className="link" onClick={() => setProblem("")}>Dismiss</button></div>}
      {view === "overview" && <LeadershipOverview clients={activeClients} me={me} onSetName={setFirstName} onOpenClient={(id, range) => { setDashClientId(id); setDashRange(range || null); setView("dashboard"); }} onGoClients={() => setView("clients")} onGoGuide={() => setView("guide")} />}
      {view === "dashboard" && <Dashboard me={me} can={can} clients={activeClients} initialClientId={dashClientId} initialRange={dashRange} onBack={() => setView("overview")} />}
      {view === "clients" && <Clients key={"clients" + navTick} me={me} can={can} clients={clients} users={users} refreshClients={refreshClients} />}
      {view === "formulas" && <Formulas />}
      {view === "guide" && <Guide me={me} can={can} onGo={setView} />}
      {view === "questions" && <Questions me={me} clients={activeClients} />}
      {view === "users" && can("users") && (REMOTE ? <People me={me} users={users} clients={clients} refreshUsers={refreshUsers} /> : <Users me={me} users={users} refreshUsers={refreshUsers} />)}
      {view === "audit" && can("audit") && <><div className="wrap" style={{ paddingBottom: 0 }}><FeedbackAdmin /></div><AuditLog /></>}
      <FeedbackWidget me={me} view={view} />
    </div>
  );
}
/* Admin is by name only. An account is Admin when its email is on this list, and only then.
   Everyone else is a Supervisor unless an Admin makes them a Manager. */
const ADMIN_EMAILS = ["__ADMIN_EMAIL__"];
const isNamedAdmin = (email) => ADMIN_EMAILS.map((e) => e.trim().toLowerCase()).includes(String(email || "").trim().toLowerCase());
// bring stored accounts in line with the list: named people are Admin, nobody else is
const alignRoles = (list) => { let changed = false; const next = (list || []).map((u) => { const want = isNamedAdmin(u.username) ? "admin" : u.role === "admin" ? "supervisor" : u.role; if (want !== u.role) { changed = true; return { ...u, role: want }; } return u; }); return { next, changed }; };
const PERMS = {
  admin: ["users", "audit", "managers", "clients:add", "clients:delete", "clients:reassign", "scopes:edit", "targets:edit", "data:entry", "period:delete"],
  // on the shared server a Manager also sees their team and can hand a client from one team member to another
  manager: [...(REMOTE ? ["users", "clients:reassign"] : []), "audit", "clients:add", "clients:delete", "scopes:edit", "targets:edit", "data:entry"],
  supervisor: ["clients:add", "clients:delete", "scopes:edit", "data:entry"],
  tl: ["clients:add", "clients:delete", "scopes:edit", "data:entry"],
  sme: ["clients:add", "clients:delete", "scopes:edit", "data:entry"],
};
const ROLE_LABEL = { admin: "Admin", manager: "Manager", supervisor: "Supervisor", tl: "Team Lead", sme: "SME" };
const roleLabel = (r) => ROLE_LABEL[r] || r || "";
const SELF_ROLES = [["supervisor", "Supervisor"], ["tl", "Team Lead (TL)"], ["sme", "SME"], ["manager", "Manager"]];

/* ============================== auth =================================== */
function Auth({ users, onUsers, onLogin }) {
  const firstRun = users.length === 0;
  const [mode, setMode] = useState(firstRun ? "create" : "login");
  const [username, setUsername] = useState(""); const [password, setPassword] = useState(""); const [confirm, setConfirm] = useState("");
  const [sa, setSa] = useState("");
  const [err, setErr] = useState(""); const [okMsg, setOkMsg] = useState("");
  const [resetUser, setResetUser] = useState(null); const [newPass, setNewPass] = useState("");
  const [busy, setBusy] = useState(false);
  const clean = (v) => v.trim().toLowerCase();
  // One simple form for everyone: Purview email, password, confirm. The very first account on a computer
  // keeps full permissions behind the scenes so users and settings can be managed later.
  const doCreate = async () => {
    setErr(""); setOkMsg("");
    const email = clean(username);
    if (!email || !password || !confirm) return setErr("Fill in all three boxes.");
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return setErr("Use your Purview email address as the username.");
    if (password.length < 6) return setErr("Use at least 6 characters for the password.");
    if (password !== confirm) return setErr("The two passwords do not match.");
    if (users.some((x) => x.username.toLowerCase() === email)) return setErr("An account with that email already exists. Sign in instead.");
    setBusy(true);
    const salt = uid();
    const u = { id: uid(), username: email, role: isNamedAdmin(email) ? "admin" : "supervisor", salt, passHash: await sha256(salt + password), createdAt: nowIso(), createdBy: "self" };
    await sSet(K.users, [...users, u]); await onUsers();
    await appendAudit({ user: u.username, role: u.role, action: "created own account" });
    setBusy(false);
    setPassword(""); setConfirm(""); setUsername(email);
    setOkMsg("Account created. Sign in with your email and password."); setMode("login");
  };
  const doLogin = async () => {
    setErr(""); setOkMsg("");
    const u = users.find((x) => x.username.toLowerCase() === clean(username));
    if (!u) return setErr("No account with that email.");
    if (await sha256(u.salt + password) !== u.passHash) return setErr("Incorrect password.");
    onLogin(u);
  };
  const startReset = () => {
    setErr(""); setOkMsg("");
    const u = users.find((x) => x.username.toLowerCase() === clean(username));
    if (!u) return setErr("Type your email first, then choose Forgot password.");
    if (!u.saHash) return setErr("Ask whoever manages Drishti on this computer to reset your password from the Users page.");
    setResetUser(u); setSa(""); setMode("reset");
  };
  const doReset = async () => {
    if (await sha256(resetUser.salt + sa.trim().toLowerCase()) !== resetUser.saHash) return setErr("Security answer doesn't match.");
    if (!newPass) return setErr("Enter a new password.");
    const salt = uid(); const nu = await Promise.all(users.map(async (x) => x.id === resetUser.id ? { ...x, salt, passHash: await sha256(salt + newPass) } : x));
    await sSet(K.users, nu); await onUsers(); await appendAudit({ user: resetUser.username, role: resetUser.role, action: "reset own password" });
    setOkMsg("Password updated. Sign in now."); setMode("login"); setSa(""); setNewPass(""); setPassword(""); setErr("");
  };
  const onEnter = (fn) => (e) => { if (e.key === "Enter") fn(); };
  return (<div className="auth-wrap"><div className="auth">
    <Brand size={26} /><div className="sub" style={{ marginTop: 10 }}>Client KPI and SLA console</div>
    {mode === "create" && <>
      <h2>Create your account</h2>
      <p style={{ color: "var(--ink-3)", fontSize: 12.5, marginTop: -8, marginBottom: 16 }}>{firstRun ? "You only do this once. After that, you just sign in." : "Set up your login once, then sign in."}</p>
      {err && <div className="msg err">{err}</div>}
      <div className="field"><label>Purview email</label><input type="email" value={username} onChange={(e) => setUsername(e.target.value)} placeholder="name@purview-healthcare.com" autoFocus autoComplete="username" onKeyDown={onEnter(doCreate)} /></div>
      <div className="field"><label>Password</label><input type="password" value={password} onChange={(e) => setPassword(e.target.value)} autoComplete="new-password" onKeyDown={onEnter(doCreate)} /></div>
      <div className="field"><label>Confirm password</label><input type="password" value={confirm} onChange={(e) => setConfirm(e.target.value)} autoComplete="new-password" onKeyDown={onEnter(doCreate)} /></div>
      <button className="btn primary" style={{ width: "100%" }} onClick={doCreate} disabled={busy}>{busy ? "Creating…" : "Create account"}</button>
      {!firstRun && <div style={{ textAlign: "center", marginTop: 14 }}><button className="link" onClick={() => { setErr(""); setMode("login"); }}>Already have an account? Sign in</button></div>}
    </>}
    {mode === "login" && <>
      <h2>Sign in</h2>{okMsg && <div className="msg ok">{okMsg}</div>}{err && <div className="msg err">{err}</div>}
      <div className="field"><label>Purview email</label><input type="email" value={username} onChange={(e) => setUsername(e.target.value)} placeholder="name@purview-healthcare.com" autoFocus={!username} autoComplete="username" onKeyDown={onEnter(doLogin)} /></div>
      <div className="field"><label>Password</label><input type="password" value={password} onChange={(e) => setPassword(e.target.value)} autoFocus={!!username} autoComplete="current-password" onKeyDown={onEnter(doLogin)} /></div>
      <button className="btn primary" style={{ width: "100%" }} onClick={doLogin}>Sign in</button>
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 14 }}>
        <button className="link" onClick={() => { setErr(""); setOkMsg(""); setPassword(""); setConfirm(""); setUsername(""); setMode("create"); }}>Create an account</button>
        <button className="link" onClick={startReset}>Forgot password?</button>
      </div>
    </>}
    {mode === "reset" && resetUser && <>
      <h2>Reset password</h2><p>Answer the security question for <b>{resetUser.username}</b>.</p>{err && <div className="msg err">{err}</div>}
      <div className="field"><label>{resetUser.sq}</label><input value={sa} onChange={(e) => setSa(e.target.value)} autoFocus /></div>
      <div className="field"><label>New password</label><input type="password" value={newPass} onChange={(e) => setNewPass(e.target.value)} /></div>
      <div className="modal-actions"><button className="btn ghost" onClick={() => { setMode("login"); setErr(""); }}>Back</button><button className="btn primary" onClick={doReset}>Update password</button></div>
    </>}
  </div></div>);
}

/* ======================= sign in on the shared server ======================= */
const EMAIL_RX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
function RemoteAuth({ note, onLogin }) {
  const [door, setDoor] = useState(null);        // null shows the two doors; "admin" or "user"
  const [mode, setMode] = useState("login");     // user door: "login" or "create"
  const [email, setEmail] = useState(""); const [password, setPassword] = useState(""); const [confirm, setConfirm] = useState("");
  const [name, setName] = useState(""); const [role, setRole] = useState(""); const [managerId, setManagerId] = useState("");
  const [mgrs, setMgrs] = useState(null); const [adminName, setAdminName] = useState("the Admin");
  const [err, setErr] = useState(note || ""); const [okMsg, setOkMsg] = useState(""); const [busy, setBusy] = useState(false); const [forgot, setForgot] = useState(false);
  const loadManagers = async () => { try { const r = await api("managers.public"); setMgrs(r.managers || []); if (r.adminName) setAdminName(r.adminName); } catch (e) { setMgrs((m) => m || []); setErr(e.message); } };
  useEffect(() => { loadManagers(); }, []);
  const go = (d) => { setDoor(d); setMode("login"); setErr(""); setOkMsg(""); setForgot(false); setPassword(""); setConfirm(""); };
  const clean = (v) => v.trim().toLowerCase();
  const signIn = async () => {
    setErr(""); setOkMsg(""); setForgot(false);
    if (!clean(email) || !password) return setErr("Enter your email and password.");
    setBusy(true);
    try { const r = await api(door === "admin" ? "adminLogin" : "login", { email: clean(email), password }); tokSet(r.token); await onLogin(r.user); }
    catch (e) { setErr(e.message); setBusy(false); }
  };
  const create = async () => {
    setErr(""); setOkMsg("");
    const em = clean(email);
    if (!name.trim()) return setErr("Enter your full name.");
    if (!EMAIL_RX.test(em)) return setErr("Enter your Purview email address.");
    if (!role) return setErr("Choose your role.");
    if (role !== "manager" && !managerId) return setErr("Choose your reporting manager. This is required.");
    if (password.length < 6) return setErr("Use at least 6 characters for the password.");
    if (password !== confirm) return setErr("The two passwords do not match.");
    setBusy(true);
    try {
      const r = await api("signup", { email: em, name: name.trim(), password, role, managerId: role === "manager" ? "" : managerId });
      const u = r.user || {};
      setOkMsg("Account created" + (u.role ? " as " + roleLabel(u.role) : "") + (u.managerName ? ", reporting to " + u.managerName : "") + ". Sign in with your email and password.");
      setMode("login"); setPassword(""); setConfirm(""); setEmail(em);
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const onEnter = (fn) => (e) => { if (e.key === "Enter" && !busy) fn(); };
  const noManagers = mgrs && mgrs.length === 0;
  return (<div className="auth-wrap"><div className={"auth" + (!door ? " doors-wide" : door === "user" && mode === "create" ? " wide" : "")}>
    <Brand size={26} /><div className="sub" style={{ marginTop: 10 }}>Client KPI and SLA console</div>
    {!door && <>
      <h2>Welcome</h2>
      {err && <div className="msg err">{err}</div>}
      <div className="doors">
        <button className="door" onClick={() => go("admin")}>
          <span className="door-ic" aria-hidden="true"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 3l7 3v5c0 4.6-3 8.4-7 10-4-1.6-7-5.4-7-10V6z" /><path d="M9 12l2 2 4-4" /></svg></span>
          <span className="door-t">Sign in as Admin</span>
          <span className="door-d">For {adminName} only. This account is already set up, so there is nothing to create.</span>
        </button>
        <button className="door" onClick={() => go("user")}>
          <span className="door-ic" aria-hidden="true"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><circle cx="9" cy="8" r="3.2" /><path d="M3.5 19c.6-3 2.8-4.8 5.5-4.8s4.9 1.8 5.5 4.8" /><circle cx="17" cy="9" r="2.5" /><path d="M16 14.4c2.3.1 4 1.7 4.5 4.6" /></svg></span>
          <span className="door-t">Sign in as User</span>
          <span className="door-d">Managers, Supervisors, Team Leads and SMEs. First time here? Create your account inside.</span>
        </button>
      </div>
    </>}
    {door === "admin" && <>
      <h2>Sign in as Admin</h2>
      {err && <div className="msg err">{err}</div>}
      <div className="field"><label>Admin email</label><input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="name@purview-healthcare.com" autoFocus autoComplete="username" onKeyDown={onEnter(signIn)} /></div>
      <div className="field"><label>Password</label><input type="password" value={password} onChange={(e) => setPassword(e.target.value)} autoComplete="current-password" onKeyDown={onEnter(signIn)} /></div>
      <button className="btn primary" style={{ width: "100%" }} onClick={signIn} disabled={busy}>{busy ? "Signing in…" : "Sign in"}</button>
      {forgot && <div className="msg info" style={{ marginTop: 12 }}>The Admin password is kept in the Google Sheet, under Extensions, Apps Script, Project Settings, Script properties, ADMIN_PASSWORD.</div>}
      <div className="auth-links"><button className="link" onClick={() => go(null)}>← Back</button><button className="link" onClick={() => setForgot(true)}>Forgot password?</button></div>
    </>}
    {door === "user" && <>
      <div className="seg auth-seg"><button className={"segbtn" + (mode === "login" ? " on" : "")} onClick={() => { setMode("login"); setErr(""); }}>Sign in</button><button className={"segbtn" + (mode === "create" ? " on" : "")} onClick={() => { setMode("create"); setErr(""); setOkMsg(""); if (!mgrs || !mgrs.length) loadManagers(); }}>Create account</button></div>
      {mode === "login" && <>
        <h2>Sign in</h2>{okMsg && <div className="msg ok">{okMsg}</div>}{err && <div className="msg err">{err}</div>}
        <div className="field"><label>Purview email</label><input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="name@purview-healthcare.com" autoFocus={!email} autoComplete="username" onKeyDown={onEnter(signIn)} /></div>
        <div className="field"><label>Password</label><input type="password" value={password} onChange={(e) => setPassword(e.target.value)} autoFocus={!!email} autoComplete="current-password" onKeyDown={onEnter(signIn)} /></div>
        <button className="btn primary" style={{ width: "100%" }} onClick={signIn} disabled={busy}>{busy ? "Signing in…" : "Sign in"}</button>
        {forgot && <div className="msg info" style={{ marginTop: 12 }}>Only {adminName} can set a new password. Ask {adminName} to reset it for you from the People page, then sign in with the new one.</div>}
        <div className="auth-links"><button className="link" onClick={() => go(null)}>← Back</button><button className="link" onClick={() => setForgot(true)}>Forgot password?</button></div>
      </>}
      {mode === "create" && <>
        <h2>Create your account</h2>
        <p className="auth-lead">You only do this once. After that, you just sign in.</p>
        {err && <div className="msg err">{err}</div>}
        <div className="field"><label>Full name</label><input value={name} onChange={(e) => setName(e.target.value)} placeholder="First and last name" autoFocus autoComplete="name" /></div>
        <div className="field"><label>Purview email</label><input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="name@purview-healthcare.com" autoComplete="username" /></div>
        <div className="two">
          <div className="field"><label>Your role</label><select value={role} onChange={(e) => setRole(e.target.value)}><option value="">Choose…</option>{SELF_ROLES.map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select></div>
          <div className="field"><label>Reporting manager <span className="req">required</span></label>
            {role === "manager"
              ? <select value="admin" disabled><option value="admin">{adminName} (Admin)</option></select>
              : <select value={managerId} onChange={(e) => setManagerId(e.target.value)} disabled={mgrs == null}><option value="">{mgrs == null ? "Loading…" : noManagers ? "No managers added yet" : "Choose…"}</option>{(mgrs || []).map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</select>}
          </div>
        </div>
        {role === "manager" && <div className="msg info">Managers report to {adminName}. Your email must already be on the managers list that {adminName} keeps. If it is not, ask {adminName} to add you first.</div>}
        {role && role !== "manager" && noManagers && <div className="msg info">No managers have been added yet. Ask {adminName} to add your manager, then come back to create your account.</div>}
        <div className="two">
          <div className="field"><label>Password</label><input type="password" value={password} onChange={(e) => setPassword(e.target.value)} autoComplete="new-password" onKeyDown={onEnter(create)} /></div>
          <div className="field"><label>Confirm password</label><input type="password" value={confirm} onChange={(e) => setConfirm(e.target.value)} autoComplete="new-password" onKeyDown={onEnter(create)} /></div>
        </div>
        <button className="btn primary" style={{ width: "100%" }} onClick={create} disabled={busy}>{busy ? "Creating…" : "Create account"}</button>
        <div className="auth-links"><button className="link" onClick={() => go(null)}>← Back</button><button className="link" onClick={() => { setMode("login"); setErr(""); }}>Already have an account? Sign in</button></div>
      </>}
    </>}
  </div></div>);
}

/* ------------------------ shared rule helpers -------------------------- */
const ruleKey = (scopeKey, critKey) => `${scopeKey}.${critKey}`;
function effectiveRule(client, scopeKey, crit) { return { ...(crit.rule || R.info()), ...(client?.rules?.[ruleKey(scopeKey, crit.key)] || {}) }; }
function deltaCls(delta, rule) { if (rule.dir === "lower") return delta < 0 ? "up" : "down"; if (rule.dir === "higher") return delta > 0 ? "up" : "down"; return "flat"; }

/* ===================== leadership overview ============================= */
// Count on-target / at-risk / missed for one period, mirroring the Dashboard tally exactly.
function overallCounts(client, period) {
  const t = { ok: 0, warn: 0, miss: 0 };
  if (!period) return t;
  const activeScopes = (client?.scopes || []).map((k) => SCOPE_BY_KEY[k]).filter(Boolean);
  for (const sc of activeScopes) {
    if (sc.criteria) for (const crit of sc.criteria) {
      const cell = period.scopes?.[sc.key]?.crit?.[crit.key]; if (!cell) continue;
      const st0 = critStatus(crit, cell, effectiveRule(client, sc.key, crit)); const st = st0 === "great" ? "ok" : st0; if (t[st] != null) t[st]++;
    }
    if (sc.kind === "tasks") for (const row of period.scopes?.[sc.key]?.tasks || []) { if (isNum(n(row.received)) && n(row.received) > 0 && isNum(n(row.achieved))) { const p = n(row.achieved) / n(row.received); t[p >= 0.9 ? "ok" : p >= 0.7 ? "warn" : "miss"]++; } }
  }
  for (const row of period.scopes?.sla?.tasks || []) { if (isNum(n(row.received)) && n(row.received) > 0 && isNum(n(row.achieved))) { const p = n(row.achieved) / n(row.received); t[p >= 0.9 ? "ok" : p >= 0.7 ? "warn" : "miss"]++; } }
  return t;
}
const bandOf = (pct) => (pct == null ? null : pct >= 90 ? "ok" : pct >= 80 ? "warn" : "miss");
// Full scorecard for a client's latest period: overall %, per-scope %, named strengths + drags, trend.
function clientScorecard(client, period, prevPeriod) {
  const activeScopes = (client?.scopes || []).map((k) => SCOPE_BY_KEY[k]).filter(Boolean);
  const scopeStats = [], strengths = [], drags = [];
  const kpi = { ok: 0, warn: 0, miss: 0 }, sla = { ok: 0, warn: 0, miss: 0 };
  const kpiByScope = [], slaByScope = [];
  const push = (arr, o) => arr.push(o);
  const pctBand = (o) => { const sc = o.ok + o.warn + o.miss; const pct = sc ? Math.round(o.ok / sc * 100) : null; return { scored: sc, pct, band: bandOf(pct) }; };
  const addTask = (short, row, kind, into, scopeKey) => {
    if (!(isNum(n(row.received)) && n(row.received) > 0 && isNum(n(row.achieved)))) return null;
    const p = n(row.achieved) / n(row.received); const st = p >= 0.9 ? "ok" : p >= 0.7 ? "warn" : "miss";
    const item = { label: `${short} · ${row.name}`, name: row.name, scopeKey: scopeKey || "sla", value: Math.round(p * 100) + "%", status: st, great: p >= 0.98, reason: row.comment, kind };
    push(st === "ok" ? strengths : drags, item); sla[st]++; if (into) into[st]++; return st;
  };
  for (const sc of activeScopes) {
    const s = { key: sc.key, name: sc.name, short: sc.short, ok: 0, warn: 0, miss: 0 };
    const sK = { ok: 0, warn: 0, miss: 0 }, sS = { ok: 0, warn: 0, miss: 0 };
    if (sc.criteria) for (const crit of sc.criteria) {
      const cell = period?.scopes?.[sc.key]?.crit?.[crit.key]; if (!cell) continue;
      const st0 = critStatus(crit, cell, effectiveRule(client, sc.key, crit)); const st = st0 === "great" ? "ok" : st0;
      if (st === "ok" || st === "warn" || st === "miss") { s[st]++; kpi[st]++; sK[st]++; push(st === "ok" ? strengths : drags, { label: `${sc.short} · ${crit.name}`, name: crit.name, scopeKey: sc.key, value: critDisplay(crit, cell), status: st, great: st0 === "great", reason: cell.reason, kind: "KPI" }); }
    }
    if (sc.kind === "tasks") for (const row of period?.scopes?.[sc.key]?.tasks || []) { const st = addTask(sc.short, row, "SLA", sS, sc.key); if (st) s[st]++; }
    const scored = s.ok + s.warn + s.miss; s.scored = scored; s.pct = scored ? Math.round(s.ok / scored * 100) : null; s.band = bandOf(s.pct);
    const kb = pctBand(sK); if (kb.scored) kpiByScope.push({ short: sc.short, ...kb });
    const sb = pctBand(sS); if (sb.scored) slaByScope.push({ short: sc.short, ...sb });
    scopeStats.push(s);
  }
  const mergedSla = { ok: 0, warn: 0, miss: 0 };
  for (const row of period?.scopes?.sla?.tasks || []) addTask("SLA", row, "SLA", mergedSla, "sla");
  { const mb = pctBand(mergedSla); if (mb.scored) slaByScope.unshift({ short: "SLA", ...mb }); }
  const c = { ok: kpi.ok + sla.ok, warn: kpi.warn + sla.warn, miss: kpi.miss + sla.miss };
  const scored = c.ok + c.warn + c.miss;
  const pct = scored ? Math.round(c.ok / scored * 100) : null; const band = bandOf(pct);
  const kpiScored = kpi.ok + kpi.warn + kpi.miss, slaScored = sla.ok + sla.warn + sla.miss;
  const kpiPct = kpiScored ? Math.round(kpi.ok / kpiScored * 100) : null, slaPct = slaScored ? Math.round(sla.ok / slaScored * 100) : null;
  let delta = null; if (prevPeriod) { const pc = overallCounts(client, prevPeriod); const ps = pc.ok + pc.warn + pc.miss; if (ps) { const pv = Math.round(pc.ok / ps * 100); if (pct != null) delta = pct - pv; } }
  const rank = { miss: 0, warn: 1, ok: 2 };
  drags.sort((a, b) => rank[a.status] - rank[b.status]);
  strengths.sort((a, b) => (b.great ? 1 : 0) - (a.great ? 1 : 0));
  const topDrags = drags.slice(0, 2);
  const topScope = scopeStats.filter((s) => s.pct != null).sort((a, b) => b.pct - a.pct)[0] || null;
  return { pct, band, scored, counts: c, kpi, sla, kpiPct, kpiBand: bandOf(kpiPct), slaPct, slaBand: bandOf(slaPct), kpiScored, slaScored, kpiByScope, slaByScope, scopeStats, strengths, drags, topDrags, topScope, delta };
}
// metrics that were red in this period and in the two before it of the same kind
function redStreak(client, list) {
  if (list.length < 3 || list.some((p) => !p)) return [];
  const names = [], active = (client?.scopes || []).map((k) => SCOPE_BY_KEY[k]).filter(Boolean);
  const red = (p, sc, crit) => { const cell = p.scopes?.[sc.key]?.crit?.[crit.key]; return !!cell && critStatus(crit, cell, effectiveRule(client, sc.key, crit)) === "miss"; };
  for (const sc of active) if (sc.criteria) for (const crit of sc.criteria) if (list.every((p) => red(p, sc, crit))) names.push(crit.name);
  const taskRed = (p, key, name) => { const r = (p.scopes?.[key]?.tasks || []).find((t) => t.name === name); if (!r) return false; const rec = n(r.received), ach = n(r.achieved); return isNum(rec) && rec > 0 && isNum(ach) && ach / rec < 0.7; };
  const slaNames = new Set(); for (const p of list) for (const r of p.scopes?.sla?.tasks || []) if (r.name) slaNames.add(r.name);
  for (const nm of slaNames) if (list.every((p) => taskRed(p, "sla", nm))) names.push(nm);
  for (const sc of active) if (sc.kind === "tasks") for (const nm of sc.tasks || []) if (list.every((p) => taskRed(p, sc.key, nm))) names.push(nm);
  return names;
}
function ScoreRing({ pct, band, size = 60 }) {
  const deg = (pct || 0) * 3.6; const inset = Math.round(size * 0.083);
  return (<div style={{ position: "relative", width: size, height: size, flex: "none", borderRadius: "50%", background: `conic-gradient(var(--${band || "info"}) ${deg}deg, var(--line-2) 0deg)` }}>
    <div style={{ position: "absolute", inset, borderRadius: "50%", background: "var(--surface)", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", boxShadow: "inset 0 1px 3px rgba(22,35,59,.08)" }}>
      <span style={{ fontSize: pct == null ? size * 0.2 : size * 0.27, fontWeight: 700, color: `var(--${band || "info"})`, letterSpacing: "-.02em", lineHeight: 1 }}>{pct == null ? "—" : pct + "%"}</span>
      {pct == null && size >= 52 && <span style={{ fontSize: Math.max(7, size * 0.12), color: "var(--ink-3)", marginTop: 2, letterSpacing: ".02em" }}>no data</span>}
    </div></div>);
}
const bandWord = (b) => (b === "ok" ? "Healthy" : b === "warn" ? "At risk" : b === "miss" ? "Below target" : "No data");
// Friendly display name from a username: strip trailing digits, split separators, title-case.
function displayName(me) {
  const set = String(me?.firstName || "").trim();
  if (set) return set.split(/\s+/)[0].replace(/^\w/, (m) => m.toUpperCase());
  // logins are Purview emails, so only ever use the part before the @
  const raw = String(me?.username || "").split("@")[0].replace(/\d+$/, "").replace(/[._-]+/g, " ").trim();
  if (!raw) return "there";
  return raw.split(/\s+/)[0].toLowerCase().replace(/^\w/, (m) => m.toUpperCase());
}
function FirstNameModal({ me, onClose, onSave }) {
  const [v, setV] = useState(String(me?.firstName || ""));
  const go = () => { onSave(v.trim()); onClose(); };
  return (
    <Modal title="What should we call you?" desc="Used for the greeting on this screen. Just your first name." onClose={onClose}>
      <div className="field"><label>First name</label><input autoFocus value={v} onChange={(e) => setV(e.target.value)} placeholder="Shreya" onKeyDown={(e) => { if (e.key === "Enter") go(); }} /></div>
      <div className="modal-actions"><button className="btn ghost" onClick={onClose}>Cancel</button><button className="btn primary" onClick={go} disabled={!v.trim()}>Save</button></div>
    </Modal>
  );
}
// A greeting that reacts to what the numbers actually say: light in tone, serious in substance.
function heroGreeting(name, { hasData, nOk, nWarn, nMiss }) {
  const h = new Date().getHours();
  const tod = h < 12 ? "Good morning" : h < 17 ? "Good afternoon" : "Good evening";
  if (!hasData) return { lead: `${tod}, `, highlight: name, tail: ". Ready to get the full picture?", sub: "Set up a client below and this whole screen comes alive.", isName: true };
  if (nMiss > 0) return { lead: "", highlight: `${nMiss} client${nMiss > 1 ? "s" : ""}`, tail: ` need${nMiss > 1 ? "" : "s"} your eye today, ${name}.`, sub: "Start at the top. The reasons are already named for you.", isName: false };
  if (nWarn > 0) return { lead: `Mostly steady, `, highlight: name, tail: `. ${nWarn} to keep an eye on.`, sub: "Nothing in the red this period. Let's keep it that way.", isName: true };
  return { lead: "Every client is on target, ", highlight: name, tail: ". Take a bow.", sub: `All ${nOk} of them green. That is a good day for the floor.`, isName: true };
}
/* isometric hero illustration — "Drishti / vision": two people gaining insight from floating dashboards under a watching eye */
function DrishtiScene() {
  const outer = [
    { x: 150, y: 62, c: "#3ECF7C" }, { x: 248, y: 88, c: "#f4cd6a" },
    { x: 150, y: 154, c: "#f2707a" }, { x: 52, y: 128, c: "#3ECF7C" },
  ];
  const inner = [
    { x: 224, y: 108, c: "#3ECF7C" }, { x: 88, y: 92, c: "#f4cd6a" }, { x: 130, y: 136, c: "#3ECF7C" },
  ];
  const Dot = ({ x, y, c, r = 4.2 }) => (<g><circle className="dsc-halo" cx={x} cy={y} r={r * 2.2} fill={c} opacity="0.2" /><circle cx={x} cy={y} r={r} fill={c} /><circle cx={x - r * 0.35} cy={y - r * 0.35} r={r * 0.32} fill="#ffffff" opacity="0.7" /></g>);
  return (
    <svg width="308" height="216" viewBox="0 0 308 216" fill="none" xmlns="http://www.w3.org/2000/svg">
      <defs>
        <linearGradient id="dg" x1="0" y1="0" x2="1" y2="1"><stop stopColor="#3ECF7C" /><stop offset="0.5" stopColor="#2CC9B0" /><stop offset="1" stopColor="#38C8E8" /></linearGradient>
        <radialGradient id="deye" cx="0.5" cy="0.5" r="0.5"><stop stopColor="#8fe9ff" stopOpacity="0.9" /><stop offset="1" stopColor="#38C8E8" stopOpacity="0" /></radialGradient>
        <clipPath id="irisClip"><circle cx="150" cy="108" r="16" /></clipPath>
      </defs>
      <ellipse cx="150" cy="108" rx="104" ry="74" fill="url(#deye)" opacity="0.42" />
      {/* outer orbit — clients circling the vision */}
      <g className="dsc-orbA">
        <ellipse cx="150" cy="108" rx="112" ry="46" stroke="rgba(255,255,255,0.22)" strokeWidth="1.2" strokeDasharray="2 6" transform="rotate(-13 150 108)" />
        {outer.map((d, i) => <Dot key={i} {...d} r="4.6" />)}
      </g>
      {/* inner orbit — counter-rotating for depth */}
      <g className="dsc-orbB">
        <ellipse cx="150" cy="108" rx="76" ry="30" stroke="rgba(255,255,255,0.15)" strokeWidth="1" strokeDasharray="2 6" transform="rotate(14 150 108)" />
        {inner.map((d, i) => <Dot key={i} {...d} r="3.4" />)}
      </g>
      {/* Drishti eye — the vision, with a score gauge for an iris and a radar sweep */}
      <g className="flo1">
        <path d="M88 108 Q150 66 212 108 Q150 150 88 108 Z" fill="#0c4640" stroke="#bdeee2" strokeWidth="2" />
        <path d="M96 106 Q150 72 204 106" stroke="#eafaf5" strokeWidth="1.4" opacity="0.35" fill="none" />
        <circle cx="150" cy="108" r="25" fill="#0d3f3a" stroke="url(#dg)" strokeWidth="3.5" />
        <g clipPath="url(#irisClip)"><g className="dsc-radar"><path d="M150 108 L150 88 A20 20 0 0 1 169 101 Z" fill="url(#dg)" opacity="0.45" /></g></g>
        <circle cx="150" cy="108" r="16" stroke="rgba(255,255,255,0.18)" strokeWidth="4" fill="none" />
        <circle cx="150" cy="108" r="16" stroke="url(#dg)" strokeWidth="4" fill="none" strokeLinecap="round" strokeDasharray="75 100" transform="rotate(-90 150 108)" />
        <circle cx="150" cy="108" r="6.5" fill="#0a3b36" /><circle cx="153" cy="105" r="2.3" fill="#eafaf5" />
        <path d="M150 74 L150 66 M190 84 L196 79 M110 84 L104 79 M200 108 L208 108 M100 108 L92 108" stroke="#7ff0d0" strokeWidth="1.6" strokeLinecap="round" opacity="0.8" />
      </g>
    </svg>
  );
}
function kudosText(c) {
  const best = c.topScope ? `, with ${c.topScope.name} at ${c.topScope.pct}%` : "";
  return `Shout-out to ${c.name}! They hit ${c.pct}% overall this period${best}. Great work, team. #PurviewProud`;
}
function KudosModal({ card, onClose }) {
  const [copied, setCopied] = useState(false);
  const text = kudosText(card);
  const copy = async () => { try { await navigator.clipboard.writeText(text); setCopied(true); setTimeout(() => setCopied(false), 1800); } catch { setCopied(false); } };
  return (
    <Overlay onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h3>Share kudos</h3><p>Copy this and post it in your team group (Bob the Builders).</p>
        <textarea readOnly value={text} style={{ width: "100%", minHeight: 96, resize: "vertical", fontSize: 13 }} />
        <div className="modal-actions"><button className="btn ghost" onClick={onClose}>Close</button><button className="btn primary" onClick={copy}>{copied ? "Copied ✓" : "Copy kudos"}</button></div>
      </div>
    </Overlay>
  );
}
function ShoutoutBanner({ stars, onShare }) {
  const top = stars[0]; const others = stars.length - 1;
  const best = top.topScope && top.topScope.pct >= 90 ? `Standout: ${top.topScope.name} at ${top.topScope.pct}%. ` : "";
  return (
    <div className="shout">
      <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
        <div style={{ fontSize: 30 }}>🏆</div>
        <div style={{ flex: 1, minWidth: 220 }}>
          <div style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".08em", opacity: .8 }}>Shout-out of the period</div>
          <div style={{ fontSize: 17, fontWeight: 700, marginTop: 2 }}>{top.name} is leading the way, {top.pct}% overall{others > 0 ? ` (and ${others} more on target)` : ""}</div>
          <div style={{ fontSize: 12.5, marginTop: 3 }}>{best}Keep it up, team.</div>
        </div>
        <button className="kbtn" onClick={() => onShare(top)}>Share kudos ↗</button>
      </div>
    </div>
  );
}
function LeadershipOverview({ clients, me, onSetName, onOpenClient, onGoClients, onGoGuide }) {
  const [cards, setCards] = useState(null);
  const [kudos, setKudos] = useState(null);
  const [nameOpen, setNameOpen] = useState(false);
  const [celebrate, setCelebrate] = useState(0);
  const [mode, setMode] = useState("latest");   // "latest", or "month" / "week" to line every client up on the same dates
  const [sel, setSel] = useState(null);         // the chosen month or week
  const [marks, setMarks] = useState(() => new Set());
  const latestRef = useRef([]);
  const choose = (m) => {
    if (m === "latest") { setMode("latest"); setSel(null); return; }
    // start on the newest month (or week) anyone has saved, otherwise the current one
    let best = null;
    for (const c of latestRef.current) for (const p of c.periods || []) if (periodType(p) === m && p.range?.to && (!best || p.range.to > best.to)) best = { from: p.range.from, to: p.range.to };
    const r = best || rangeFor(m, new Date());
    setSel({ kind: m, ...r, label: periodLabel(m, r.from, r.to) }); setMode(m);
  };
  const step = (k) => { if (!sel) return; const r = stepRange(sel.kind, sel.from, k); setSel({ kind: sel.kind, ...r, label: periodLabel(sel.kind, r.from, r.to) }); };
  useEffect(() => { let live = true; (async () => {
    if (mode !== "latest" && !sel) return;
    setCards(null);
    const data = await loadOverview(mode === "latest" ? {} : { type: sel.kind, from: sel.from, to: sel.to }, clients);
    if (!live) return;
    if (mode === "latest") latestRef.current = data;
    const mk = new Set(); for (const c of [...latestRef.current, ...data]) for (const p of c.periods || []) mk.add(periodType(p) + "|" + p.range?.from + "|" + p.range?.to);
    setMarks(mk);
    const owners = Object.fromEntries((clients || []).map((c) => [c.id, c.ownerName]));
    const out = [];
    for (const full of data) {
      const base = { id: full.id, ownerName: full.ownerName || owners[full.id] || "", name: full.name, scopes: full.scopes || [] };
      const periods = full.periods || [];
      const idx = mode === "latest" ? (periods.length ? 0 : -1) : periods.findIndex((p) => periodType(p) === sel.kind && p.range?.from === sel.from && p.range?.to === sel.to);
      if (idx < 0) { out.push({ ...base, empty: true, noPeriodForRange: mode !== "latest", rangeUnit: mode !== "latest" ? TYPE_UNIT[sel.kind] : "" }); continue; }
      const period = periods[idx], prev = prevSameType(periods, idx), prev2 = prevSameType(periods, idx, 2), unit = TYPE_UNIT[periodType(period)];
      const sc = clientScorecard(full, period, prev);
      out.push({ ...base, period, periodLabel: period.label, cmpText: unit === "period" ? "vs the previous period" : "vs last " + unit, streak: redStreak(full, [period, prev, prev2]), streakUnit: unit, ...sc });
    }
    setCards(out); if (out.some((c) => c.band === "ok")) setCelebrate((x) => x + 1);
  })(); return () => { live = false; }; }, [clients, mode, sel?.kind, sel?.from]);

  if (cards == null) return <div className="wrap"><div style={{ color: "var(--ink-3)", padding: 40 }}>Loading client scores…</div></div>;

  const scoredCards = cards.filter((c) => !c.empty && c.pct != null);
  const stars = scoredCards.filter((c) => c.band === "ok").sort((a, b) => b.pct - a.pct);
  const rank = { miss: 0, warn: 1, ok: 2 };
  const sorted = [...cards].sort((a, b) => {
    if (a.empty && !b.empty) return 1; if (!a.empty && b.empty) return -1; if (a.empty && b.empty) return a.name.localeCompare(b.name);
    if (a.pct == null && b.pct != null) return 1; if (a.pct != null && b.pct == null) return -1;
    const r = (rank[a.band] ?? 3) - (rank[b.band] ?? 3); return r !== 0 ? r : a.pct - b.pct;
  });
  const tot = scoredCards.reduce((s, c) => ({ ok: s.ok + c.counts.ok, warn: s.warn + c.counts.warn, miss: s.miss + c.counts.miss }), { ok: 0, warn: 0, miss: 0 });
  const totScored = tot.ok + tot.warn + tot.miss; const coPct = totScored ? Math.round(tot.ok / totScored * 100) : null; const coBand = bandOf(coPct);
  const kTot = scoredCards.reduce((s, c) => ({ ok: s.ok + c.kpi.ok, n: s.n + c.kpiScored }), { ok: 0, n: 0 });
  const sTot = scoredCards.reduce((s, c) => ({ ok: s.ok + c.sla.ok, n: s.n + c.slaScored }), { ok: 0, n: 0 });
  const coKpi = kTot.n ? Math.round(kTot.ok / kTot.n * 100) : null, coSla = sTot.n ? Math.round(sTot.ok / sTot.n * 100) : null;
  const nOk = scoredCards.filter((c) => c.band === "ok").length, nWarn = scoredCards.filter((c) => c.band === "warn").length, nMiss = scoredCards.filter((c) => c.band === "miss").length;
  const scopeCount = new Set(cards.flatMap((c) => c.scopes)).size;

  return (
    <div className="wrap">
      <div className="pagehead">
        <div><h1>Leadership overview</h1><div className="sub">{mode === "latest" ? "Each client's most recent period, compared with the one before it of the same kind. Switch to Monthly or Weekly to line every client up on the same dates." : <>Every client for <b>{sel?.label}</b>, compared with the {TYPE_UNIT[sel?.kind]} before. Use the arrows to go back and forth.</>}</div></div>
        <div className="row ov-ctl">
          <div className="seg">{[["latest", "Latest"], ["month", "Monthly"], ["week", "Weekly"]].map(([k, l]) => <button key={k} className={"segbtn" + (mode === k ? " on" : "")} onClick={() => choose(k)}>{l}</button>)}</div>
          {mode !== "latest" && sel && <div className="ov-step">
            <button className="calnav ov-arrow" aria-label={"Previous " + TYPE_UNIT[sel.kind]} onClick={() => step(-1)}>‹</button>
            <PeriodPicker value={sel} types={[sel.kind]} marks={marks} align="right" onSelect={(r) => setSel({ kind: r.kind, from: r.from, to: r.to, label: r.label })} />
            <button className="calnav ov-arrow" aria-label={"Next " + TYPE_UNIT[sel.kind]} onClick={() => step(1)}>›</button>
          </div>}
        </div>
      </div>

      <div className="ov-hero">
        <span className="runline" />
        <svg className="hexdeco" width="180" height="180" viewBox="0 0 180 180" fill="none" aria-hidden="true"><path d="M90 8 L150 43 L150 113 L90 148 L30 113 L30 43 Z" stroke="#7ff0d0" strokeWidth="2" /><path d="M120 60 L156 81 L156 123 L120 144 L84 123 L84 81 Z" stroke="#8fe6f5" strokeWidth="2" /></svg>
        <div className="ov-illo" aria-hidden="true"><DrishtiScene /></div>
        {(() => {
          const g = heroGreeting(displayName(me), { hasData: scoredCards.length > 0, nOk, nWarn, nMiss });
          return (<div style={{ maxWidth: 470, marginBottom: 16 }}>
            <div style={{ fontSize: 11, color: "rgba(234,250,245,.72)", textTransform: "uppercase", letterSpacing: ".08em", fontWeight: 600 }}>Purview Healthcare · Leadership view</div>
            <div style={{ fontSize: 23, fontWeight: 700, letterSpacing: "-.02em", marginTop: 5, lineHeight: 1.18 }}>{g.lead}<span style={{ color: "#8fe9d6", cursor: g.isName && onSetName ? "pointer" : "inherit", borderBottom: g.isName && onSetName ? "1px dashed rgba(143,233,214,.45)" : "none" }} title={g.isName && onSetName ? "Click to set your first name" : undefined} onClick={() => { if (g.isName && onSetName) setNameOpen(true); }}>{g.highlight}</span>{g.tail}</div>
            <div style={{ fontSize: 12.5, color: "rgba(234,250,245,.76)", marginTop: 6 }}>{g.sub}</div>
          </div>);
        })()}
        {scoredCards.length === 0 ? (
          <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap", background: "rgba(255,255,255,.06)", border: "1px solid rgba(255,255,255,.12)", borderRadius: 16, padding: "16px 18px" }}>
            <ScoreRing pct={null} band="info" size={62} />
            <div style={{ flex: 1, minWidth: 220 }}>
              <div style={{ fontSize: 14, fontWeight: 600 }}>{clients.length ? "No client scores yet" : "Let's add your first client"}</div>
              <div style={{ fontSize: 12, color: "rgba(234,250,245,.72)", marginTop: 3, lineHeight: 1.5, maxWidth: 440 }}>{clients.length ? "Open a client below, pick a duration, and upload its reports. This dial will then show the share of every client's KPI and SLA metrics meeting target." : "Name the client, tick its scopes of work, pick a duration and drop in its reports. This dial then shows how every client is doing, in one place."}</div>
              {!clients.length && onGoClients && <div className="row" style={{ marginTop: 12, gap: 8 }}><button className="btn sm primary" onClick={onGoClients}>Add your first client →</button>{onGoGuide && <button className="btn sm" onClick={onGoGuide}>New here? Read the guide</button>}</div>}
            </div>
          </div>
        ) : (
        <div style={{ display: "flex", alignItems: "center", gap: 26, flexWrap: "wrap", width: "100%" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
            <div className="ov-ring-glow"><ScoreRing pct={coPct} band={coBand} size={78} /></div>
            <div style={{ maxWidth: 260 }}>
              <div style={{ fontSize: 12.5, fontWeight: 600, color: "#eafaf5" }}>Overall health, all clients</div>
              <div style={{ fontSize: 10.5, color: "rgba(234,250,245,.66)", marginTop: 2, lineHeight: 1.45 }}>Share of every client's KPI and SLA metrics that are meeting target, for the selected duration.</div>
              <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
                <span style={{ fontSize: 11.5, fontWeight: 600, padding: "3px 11px", borderRadius: 999, background: "rgba(255,255,255,.14)", color: "#eafaf5" }}>KPI {coKpi == null ? "—" : coKpi + "%"}</span>
                <span style={{ fontSize: 11.5, fontWeight: 600, padding: "3px 11px", borderRadius: 999, background: "rgba(255,255,255,.14)", color: "#eafaf5" }}>SLA {coSla == null ? "—" : coSla + "%"}</span>
              </div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            {[["healthy", nOk, "#5fe3b0"], ["at risk", nWarn, "#f4cd6a"], ["below target", nMiss, "#f2a0a0"]].map(([lab, val, col]) => (
              <div key={lab} style={{ background: "rgba(255,255,255,.08)", border: "1px solid rgba(255,255,255,.12)", borderRadius: 14, padding: "9px 15px", minWidth: 78 }}>
                <div style={{ fontSize: 20, fontWeight: 700, color: col }}>{val}</div>
                <div style={{ fontSize: 10.5, color: "rgba(234,250,245,.7)", textTransform: "uppercase", letterSpacing: ".04em" }}>{lab}</div>
              </div>
            ))}
          </div>
          <div style={{ flex: 1, minWidth: 140, textAlign: "right", fontSize: 12, color: "rgba(234,250,245,.62)" }}>{scoredCards.length} client{scoredCards.length !== 1 ? "s" : ""} scored · {scopeCount} scope{scopeCount !== 1 ? "s" : ""} tracked</div>
        </div>
        )}
      </div>

      {stars.length > 0 && <ShoutoutBanner stars={stars} onShare={(c) => setKudos(c)} />}

      <div className="ov-grid">
        {sorted.map((c) => <OverviewClientCard key={c.id} c={c} showOwner={REMOTE && (me.role === "admin" || me.role === "manager")} onOpen={() => onOpenClient(c.id, mode !== "latest" ? sel : null)} onShare={() => setKudos(c)} />)}
      </div>
      <Confetti trigger={celebrate} />
      {kudos && <KudosModal card={kudos} onClose={() => setKudos(null)} />}
      {nameOpen && <FirstNameModal me={me} onClose={() => setNameOpen(false)} onSave={onSetName} />}
    </div>
  );
}
function HexMotif({ color }) {
  return (<svg className="ov-hex" viewBox="0 0 150 150" fill="none" aria-hidden="true"><path d="M75 6 L138 42 L138 108 L75 144 L12 108 L12 42 Z" stroke={color} strokeWidth="3" /><path d="M75 34 L110 54 L110 96 L75 116 L40 96 L40 54 Z" stroke={color} strokeWidth="2.5" /><path d="M75 62 L94 73 L94 95 L75 106 L56 95 L56 73 Z" fill={color} opacity="0.5" /></svg>);
}
function OverviewClientCard({ c, onOpen, onShare, showOwner }) {
  const owner = showOwner && c.ownerName ? <span className="ownertag" title="Who maintains this client">{c.ownerName}</span> : null;
  if (c.empty) return (
    <div className="ov-card info" onClick={onOpen} style={{ cursor: "pointer", display: "flex", alignItems: "center", gap: 16 }}>
      <span className="ov-topline" style={{ backgroundImage: "linear-gradient(90deg,transparent,rgba(255,255,255,.95) 50%,transparent),linear-gradient(90deg,#3ECF7C,#2CC9B0,#38C8E8,#2CC9B0,#3ECF7C)" }} />
      <HexMotif color="#2CC9B0" />
      <ScoreRing pct={null} band="info" size={52} />
      <div style={{ flex: 1 }}><div style={{ fontSize: 15, fontWeight: 600, display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>{c.name}{owner}</div><div style={{ fontSize: 12, color: "var(--ink-3)", marginTop: 3 }}>{c.noPeriodForRange ? `Nothing uploaded for this ${c.rangeUnit || "period"} yet. Open the client to add it.` : "No period recorded yet. Open the client to upload reports."}</div></div>
      <span style={{ fontSize: 12, color: "var(--steel)", fontWeight: 500, whiteSpace: "nowrap" }}>Open client →</span>
    </div>
  );
  const hasData = c.scored > 0;
  const LINE = { ok: "linear-gradient(90deg,#1f9d6b,#3ECF7C,#2CC9B0,#3ECF7C,#1f9d6b)", warn: "linear-gradient(90deg,#c07d10,#e6a93a,#f0c265,#e6a93a,#c07d10)", miss: "linear-gradient(90deg,#cf3d3d,#e26a6a,#f0999a,#e26a6a,#cf3d3d)", info: "linear-gradient(90deg,#3ECF7C,#2CC9B0,#38C8E8,#2CC9B0,#3ECF7C)" };
  const isStar = c.band === "ok";
  return (
    <div className={"ov-card " + (c.band || "info") + (isStar ? " star" : "")}>
      <span className="ov-topline" style={{ backgroundImage: "linear-gradient(90deg,transparent,rgba(255,255,255,.95) 50%,transparent)," + LINE[c.band || "info"] }} />
      <HexMotif color={c.band === "ok" ? "#1f9d6b" : c.band === "warn" ? "#c07d10" : c.band === "miss" ? "#cf3d3d" : "#2CC9B0"} />
      {isStar && <div className="starribbon">★ STAR PERFORMER</div>}
      <div style={{ display: "flex", alignItems: hasData ? "center" : "flex-start", gap: 16, flexWrap: "wrap", cursor: "pointer" }} onClick={onOpen}>
        <div className={c.band === "ok" ? "ov-ring-glow" : ""} style={hasData ? undefined : { marginTop: 2 }}><ScoreRing pct={c.pct} band={c.band} size={60} /></div>
        <div style={{ flex: 1, minWidth: 180 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
            <span style={{ fontSize: 15, fontWeight: 600 }}>{c.name}</span>
            <span className={"chip " + (c.band || "info")}>{bandWord(c.band)}</span>
            {owner}
            {c.periodLabel && <span style={{ fontSize: 11, color: "var(--ink-3)" }}>· {c.periodLabel}</span>}
            {c.delta != null && c.delta !== 0 && <span style={{ fontSize: 11, fontWeight: 500, color: c.delta > 0 ? "var(--ok)" : "var(--miss)" }}>{c.delta > 0 ? "▲" : "▼"} {Math.abs(c.delta)}% {c.cmpText || "vs last period"}</span>}
          </div>
          {isStar && c.topScope && <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap", alignItems: "center" }}>
            <span className="bestchip">🏅 Best in {c.topScope.name} · {c.topScope.pct}%</span>
            {onShare && <button className="link" style={{ fontWeight: 600 }} onClick={(e) => { e.stopPropagation(); onShare(); }}>Share kudos ↗</button>}
          </div>}
          {!hasData && (() => {
            const rows = (c.scopes || []).map((s) => {
              const sc = SCOPE_BY_KEY[s]; const col = scopeColor(s);
              const need = scopeReportKeys(s, c.scopes); const up = need.filter((rk) => c.period?.reports?.[rk]).length;
              const critN = sc?.criteria?.length || 0, taskN = sc?.tasks?.length || 0;
              const ps = c.period?.scopes?.[s];
              const started = up > 0 || Object.values(ps?.crit || {}).some((v) => v && v.value != null && v.value !== "") || (ps?.tasks || []).some((t) => t.received || t.achieved);
              return { s, sc, col, need: need.length, up, critN, taskN, started };
            });
            const done = rows.filter((r) => r.started).length; const total = rows.length || 1; const prog = Math.round(done / total * 100);
            return (<div style={{ marginTop: 10 }}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(168px,1fr))", gap: 7 }}>
                {rows.map((r) => (
                  <div key={r.s} style={{ display: "flex", gap: 9, alignItems: "center", padding: "7px 9px", borderRadius: 11, background: "rgba(255,255,255,.7)", border: `1px solid ${r.col}44` }}>
                    <span style={{ width: 28, height: 28, flex: "none", borderRadius: 8, background: r.col + "22", color: r.col, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 9.5 }}>{scopeBadge(r.s)}</span>
                    <div style={{ minWidth: 0 }}>
                      <div style={{ fontSize: 11.5, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{r.sc?.name || r.s}</div>
                      <div style={{ fontSize: 10, color: "var(--ink-3)" }}>{r.need ? `${r.up}/${r.need} reports` : "manual entry"}{r.critN ? ` · ${r.critN} KPI` : ""}{r.taskN ? ` · ${r.taskN} SLA` : ""}</div>
                    </div>
                    <span style={{ marginLeft: "auto", width: 8, height: 8, borderRadius: "50%", background: r.started ? "var(--ok)" : "var(--line)", flex: "none" }} title={r.started ? "Started" : "Not started"} />
                  </div>
                ))}
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 10 }}>
                <div style={{ flex: 1, height: 7, borderRadius: 4, background: "rgba(22,35,59,.08)", overflow: "hidden" }}><span style={{ display: "block", height: "100%", width: prog + "%", background: "var(--grad)", transition: "width .5s var(--ease)" }} /></div>
                <span style={{ fontSize: 11, color: "var(--ink-2)", fontWeight: 600, whiteSpace: "nowrap" }}>Setup {prog}% · {done}/{rows.length} scopes started</span>
              </div>
              <div style={{ display: "flex", gap: 6, marginTop: 10, flexWrap: "wrap", alignItems: "center" }}>
                <span style={{ fontSize: 10, color: "var(--ink-3)", textTransform: "uppercase", letterSpacing: ".04em", fontWeight: 600, marginRight: 2 }}>Will show here</span>
                {["Overall KPI", "Overall SLA", "KPI by scope", "SLA by scope", "Holding it up", "Dragging it down"].map((t) => <span key={t} style={{ fontSize: 10.5, padding: "3px 10px", borderRadius: 999, border: "1px dashed var(--line)", color: "var(--ink-3)", background: "rgba(255,255,255,.5)" }}>{t}</span>)}
              </div>
            </div>);
          })()}
          {hasData && <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap", alignItems: "center" }}>
            <span style={{ fontSize: 11.5, fontWeight: 600, padding: "3px 10px", borderRadius: 7, background: `var(--${c.kpiBand || "info"}-bg)`, color: `var(--${c.kpiBand || "info"})` }}>Overall KPI {c.kpiPct == null ? "—" : c.kpiPct + "%"}</span>
            <span style={{ fontSize: 11.5, fontWeight: 600, padding: "3px 10px", borderRadius: 7, background: `var(--${c.slaBand || "info"}-bg)`, color: `var(--${c.slaBand || "info"})` }}>Overall SLA {c.slaPct == null ? "—" : c.slaPct + "%"}</span>
          </div>}
          {c.kpiByScope && c.kpiByScope.length > 0 && <div style={{ display: "flex", gap: 5, marginTop: 8, flexWrap: "wrap", alignItems: "center" }}>
            <span style={{ fontSize: 10, fontWeight: 600, color: "var(--ink-3)", textTransform: "uppercase", letterSpacing: ".04em", marginRight: 2 }}>KPI by scope</span>
            {c.kpiByScope.map((s, i) => <span key={i} style={{ fontSize: 10.5, padding: "3px 9px", borderRadius: 6, background: `var(--${s.band}-bg)`, color: `var(--${s.band})` }}>{s.short} {s.pct}%</span>)}
          </div>}
          {c.slaByScope && c.slaByScope.length > 0 && <div style={{ display: "flex", gap: 5, marginTop: 6, flexWrap: "wrap", alignItems: "center" }}>
            <span style={{ fontSize: 10, fontWeight: 600, color: "var(--ink-3)", textTransform: "uppercase", letterSpacing: ".04em", marginRight: 2 }}>SLA by scope</span>
            {c.slaByScope.map((s, i) => <span key={i} style={{ fontSize: 10.5, padding: "3px 9px", borderRadius: 6, background: `var(--${s.band}-bg)`, color: `var(--${s.band})` }}>{s.short} {s.pct}%</span>)}
          </div>}
          {c.streak && c.streak.length > 0 && <div className="ov-streak" title={"Red in each of the last 3 " + c.streakUnit + "s: " + c.streak.join(", ")}>
            <span className="ov-streak-ic" aria-hidden="true">!</span>Red 3 {c.streakUnit}s in a row: <b>{c.streak.slice(0, 2).join(", ")}</b>{c.streak.length > 2 ? ` and ${c.streak.length - 2} more` : ""}
          </div>}
        </div>
        {hasData ? <span style={{ fontSize: 12, color: "var(--steel)", fontWeight: 500, whiteSpace: "nowrap" }}>Open client →</span>
          : <button className="btn primary sm" style={{ alignSelf: "flex-start", whiteSpace: "nowrap" }} onClick={(e) => { e.stopPropagation(); onOpen(); }}>Set up {c.name} →</button>}
      </div>
      {hasData && <MetricSplit strengths={c.strengths} drags={c.drags} />}
    </div>
  );
}

/* What is holding a client up and what is dragging it down, kept short:
   always the top three on each side (worst or best first, one per scope before repeating a scope).
   Everything else sits in a dropdown, grouped by scope, for anyone who wants to read it all. */
const SPLIT_TOP = 3;
function pickSpread(list, n, tierOf) {
  const out = [];
  for (const t of [...new Set(list.map(tierOf))]) {
    if (out.length >= n) break;
    const tier = list.filter((x) => tierOf(x) === t), seen = new Set(out.map((x) => x.scopeKey)), picked = [];
    for (const x of tier) if (out.length + picked.length < n && !seen.has(x.scopeKey)) { picked.push(x); seen.add(x.scopeKey); }
    for (const x of tier) if (out.length + picked.length < n && !picked.includes(x)) picked.push(x);
    out.push(...tier.filter((x) => picked.includes(x)));
  }
  return out;
}
function groupByScope(list) {
  const groups = [];
  for (const x of list) { let g = groups.find((y) => y.key === x.scopeKey); if (!g) groups.push((g = { key: x.scopeKey, items: [] })); g.items.push(x); }
  return groups;
}
const scopeFullName = (k) => SCOPE_BY_KEY[k]?.name || (k === "sla" ? "Service Level Agreement" : k);
const ScopeTag = ({ k }) => <span className="ov-stag" style={{ color: scopeColor(k), background: scopeColor(k) + "1c" }}>{scopeBadge(k)}</span>;
// the dropdown is drawn on the page body, so the card's rounded edges never cut it off
function MoreDropdown({ title, items, good }) {
  const [open, setOpen] = useState(false);
  const [pos, setPos] = useState(null);
  const btnRef = useRef(null), panelRef = useRef(null);
  const place = () => {
    const btn = btnRef.current; if (!btn) return;
    const side = btn.closest(".ov-side") || btn;
    const r = side.getBoundingClientRect(), br = btn.getBoundingClientRect();
    const vw = window.innerWidth, vh = window.innerHeight;
    const width = Math.min(Math.max(r.width, 290), 460, vw - 16);
    const left = Math.max(8, Math.min(r.left, vw - width - 8));
    const below = vh - br.bottom - 14, above = br.top - 14;
    const up = below < 240 && above > below;
    const maxH = Math.max(170, Math.min(400, up ? above : below));
    setPos(up ? { left, width, bottom: vh - br.top + 6, maxH } : { left, width, top: br.bottom + 6, maxH });
  };
  useEffect(() => {
    if (!open) return;
    place();
    const onDown = (e) => { if (panelRef.current?.contains(e.target) || btnRef.current?.contains(e.target)) return; setOpen(false); };
    const onKey = (e) => { if (e.key === "Escape") { setOpen(false); btnRef.current?.focus(); } };
    const onMove = () => place();
    document.addEventListener("mousedown", onDown); document.addEventListener("keydown", onKey);
    window.addEventListener("scroll", onMove, true); window.addEventListener("resize", onMove);
    return () => { document.removeEventListener("mousedown", onDown); document.removeEventListener("keydown", onKey); window.removeEventListener("scroll", onMove, true); window.removeEventListener("resize", onMove); };
  }, [open]);
  const groups = groupByScope(items);
  const panel = open && pos ? (
    <div ref={panelRef} className="ov-dd" role="dialog" aria-label={title} style={{ left: pos.left, width: pos.width, top: pos.top, bottom: pos.bottom, maxHeight: pos.maxH }}>
      <div className={"ov-dd-h " + (good ? "ok" : "miss")}><span>{title}</span><button className="ov-dd-x" aria-label="Close" onClick={() => setOpen(false)}>×</button></div>
      <div className="ov-dd-body">
        {groups.map((g) => (
          <div key={g.key} className="ov-dd-grp">
            <div className="ov-dd-gh"><ScopeTag k={g.key} /><span>{scopeFullName(g.key)}</span><span className="ov-dd-gn">{g.items.length}</span></div>
            {g.items.map((x, i) => (
              <div key={i} className="ov-dd-row" title={x.reason ? x.name + ": " + x.reason : x.name}>
                <span className="ov-mname">{x.name}</span>
                <span className={"ov-mval " + (good ? "ok" : x.status)}>{x.value}</span>
                {x.reason && <span className="ov-why">{x.reason}</span>}
              </div>
            ))}
          </div>
        ))}
      </div>
    </div>
  ) : null;
  return (<>
    <button ref={btnRef} className={"ov-dd-btn" + (open ? " on" : "")} aria-haspopup="dialog" aria-expanded={open} onClick={(e) => { e.stopPropagation(); setOpen((o) => !o); }}>
      <span>{items.length} more</span>
      <svg width="11" height="11" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5 6 8l3.5-3.5" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" /></svg>
    </button>
    {panel && (typeof document !== "undefined" && document.body ? createPortal(panel, document.body) : panel)}
  </>);
}
function SplitSide({ kind, items }) {
  const good = kind === "up";
  const top = good ? pickSpread(items, SPLIT_TOP, (x) => (x.great ? 0 : 1)) : pickSpread(items, SPLIT_TOP, (x) => x.status);
  const rest = items.filter((x) => !top.includes(x));
  const reds = items.filter((x) => x.status === "miss").length, ambers = items.filter((x) => x.status === "warn").length;
  return (
    <div className="ov-side">
      <div className={"ov-colh " + (good ? "ok" : "miss")}>
        <span>{good ? "Holding it up" : "Dragging it down"}</span>
        {items.length > 0 && <span className="n">{items.length}</span>}
        {!good && items.length > 0 && <span className="ov-mix">{[reds && reds + " red", ambers && ambers + " amber"].filter(Boolean).join(" · ")}</span>}
      </div>
      {items.length === 0 ? <div className={"ov-none" + (good ? "" : " ok")}>{good ? "Nothing on target yet." : "Nothing flagged. All metrics on target."}</div> : <>
        {top.map((x, i) => (
          <div key={i} className="ov-mrow" title={x.reason ? x.name + ": " + x.reason : x.name}>
            <ScopeTag k={x.scopeKey} />
            <span className="ov-mname">{x.name}</span>
            <span className={"ov-mval " + (good ? "ok" : x.status)}>{x.value}</span>
            {x.reason && <span className="ov-why">{x.reason}</span>}
          </div>
        ))}
        {rest.length > 0 && <div className="ov-more"><MoreDropdown title={(good ? "Also holding it up" : "Also dragging it down") + " · " + rest.length} items={rest} good={good} /></div>}
      </>}
    </div>
  );
}
function MetricSplit({ strengths, drags }) {
  return (<div className="ov-split"><SplitSide kind="up" items={strengths || []} /><SplitSide kind="down" items={drags || []} /></div>);
}

/* ============================ dashboard ================================ */
function Dashboard({ me, can, clients, initialClientId, initialRange, onBack }) {
  const [selId, setSelId] = useState(initialClientId || clients[0]?.id || null);
  const [client, setClient] = useState(null);
  const [pIdx, setPIdx] = useState(0);
  const [editing, setEditing] = useState(null);
  const [info, setInfo] = useState(null);
  const [delPeriod, setDelPeriod] = useState(false);
  const [tab, setTab] = useState("period");
  useEffect(() => { if (initialClientId) setSelId(initialClientId); }, [initialClientId]);
  useEffect(() => { if (clients.length && !selId) setSelId(clients[0].id); }, [clients]);
  useEffect(() => { (async () => {
    const c = selId ? await sGet(clientKey(selId)) : null; if (c) c.periods = sortPeriods(c.periods);
    setClient(c);
    // opened from the Overview for a chosen month or week: start on that one
    const i = initialRange && c ? c.periods.findIndex((p) => p.range?.from === initialRange.from && p.range?.to === initialRange.to) : -1;
    setPIdx(i >= 0 ? i : 0);
  })(); }, [selId]);

  const period = client?.periods?.[pIdx] || null;
  const prevIdx = client && period ? client.periods.indexOf(prevSameType(client.periods, pIdx)) : -1;
  const prev = prevIdx >= 0 ? client.periods[prevIdx] : null;
  // this period and the one it is compared with must be loaded in full
  useEffect(() => {
    if (!client) return; let live = true;
    const need = [period, prev].filter((p) => p && p._stub);
    if (!need.length) return;
    (async () => {
      const got = await Promise.all(need.map((p) => loadFullPeriod(client.id, p.id)));
      if (!live) return;
      setClient((c) => c && c.id === client.id ? { ...c, periods: c.periods.map((x) => got.find((g) => g && g.id === x.id) || x) } : c);
    })();
    return () => { live = false; };
  }, [client?.id, period?.id, period?._stub, prev?.id, prev?._stub]);
  const activeScopes = (client?.scopes || []).map((k) => SCOPE_BY_KEY[k]).filter(Boolean);

  const tally = useMemo(() => {
    const t = { ok: 0, warn: 0, miss: 0 };
    if (!period) return t;
    for (const sc of activeScopes) {
      if (sc.criteria) for (const crit of sc.criteria) {
        const cell = period.scopes?.[sc.key]?.crit?.[crit.key]; if (!cell) continue;
        const st0 = critStatus(crit, cell, effectiveRule(client, sc.key, crit)); const st = st0 === "great" ? "ok" : st0; if (t[st] != null) t[st]++;
      }
      if (sc.kind === "tasks") {
        for (const row of period.scopes?.[sc.key]?.tasks || []) { if (isNum(n(row.received)) && n(row.received) > 0 && isNum(n(row.achieved))) { const p = n(row.achieved) / n(row.received); t[p >= 0.9 ? "ok" : p >= 0.7 ? "warn" : "miss"]++; } }
      }
    }
    for (const row of period.scopes?.sla?.tasks || []) { if (isNum(n(row.received)) && n(row.received) > 0 && isNum(n(row.achieved))) { const p = n(row.achieved) / n(row.received); t[p >= 0.9 ? "ok" : p >= 0.7 ? "warn" : "miss"]++; } }
    return t;
  }, [period, client]);
  const scored = tally.ok + tally.warn + tally.miss;

  if (!clients.length) return <div className="wrap"><div className="empty"><p>No clients yet. Add one from the <b>Clients</b> tab, then enter data.</p></div></div>;
  return (
    <div className="wrap">
      {onBack && <button className="link" style={{ marginBottom: 12, display: "inline-flex", alignItems: "center", gap: 5, textDecoration: "none", fontWeight: 500 }} onClick={onBack}>← Back to overview</button>}
      <div className="pagehead">
        <div><h1>{client?.name || "Client"}</h1><div className="sub">{client ? <span className="chips" style={{ display: "inline-flex" }}>{(client.scopes || []).map((k) => <span key={k} className="pill">{SCOPE_BY_KEY[k]?.short}</span>)}</span> : ""}</div></div>
        <div className="row">
          <select value={selId || ""} onChange={(e) => setSelId(e.target.value)}>{clients.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</select>
          {tab === "period" && client?.periods?.length > 0 && <PeriodSelect periods={client.periods} pIdx={pIdx} onIndex={setPIdx} />}
          <ViewSwitch tab={tab} onTab={setTab} />
        </div>
      </div>
      {tab === "history" && client && <ClientHistory client={client} onOpenPeriod={(id) => { const i = client.periods.findIndex((p) => p.id === id); if (i >= 0) { setPIdx(i); setTab("period"); } }} />}
      {tab === "period" && !period && <div className="card" style={{ color: "var(--ink-3)" }}>Nothing recorded for this client yet. Open it under <b>Clients</b> to upload reports and add a period.</div>}
      {tab === "period" && period && (period._stub || prev?._stub) && <div className="card" style={{ color: "var(--ink-3)" }}>Loading {period.label}…</div>}
      {tab === "period" && period && !period._stub && !prev?._stub && <>
        <div className="strip">
          {scored > 0 && (() => { const pct = Math.round(tally.ok / scored * 100); const band = pct >= 90 ? "ok" : pct >= 80 ? "warn" : "miss"; return (<div style={{ display: "flex", alignItems: "center", gap: 16, marginBottom: 10 }}><div style={{ position: "relative", width: 76, height: 76, flex: "none", borderRadius: "50%", background: `conic-gradient(var(--${band}) ${pct * 3.6}deg, var(--line-2) 0deg)` }}><div style={{ position: "absolute", inset: 6, borderRadius: "50%", background: "var(--surface)", display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "inset 0 1px 3px rgba(22,35,59,.08)" }}><span style={{ fontSize: 21, fontWeight: 700, color: `var(--${band})`, letterSpacing: "-.02em" }}>{pct}%</span></div></div><div><div style={{ fontSize: 13.5, fontWeight: 600, letterSpacing: "-.01em" }}>Overall KPI + SLA score</div><div style={{ fontSize: 11.5, color: "var(--ink-3)", marginTop: 2 }}>{band === "ok" ? "On target (≥90%)" : band === "warn" ? "At risk (80 to 89%)" : "Below target (under 80%), see flagged metrics below for reasons"}</div></div></div>); })()}
          <div className="headline">{scored > 0 ? <><b>{tally.ok}</b> of <b>{scored}</b> scored metrics on target</> : "Awaiting scored metrics"}</div>
          <div className="bar">{scored > 0 && <><span style={{ width: `${(tally.ok / scored) * 100}%`, background: "var(--ok)" }} /><span style={{ width: `${(tally.warn / scored) * 100}%`, background: "var(--warn)" }} /><span style={{ width: `${(tally.miss / scored) * 100}%`, background: "var(--miss)" }} /></>}</div>
          <div className="tallies"><div className="tally ok"><span className="n">{tally.ok}</span> on target</div><div className="tally warn"><span className="n">{tally.warn}</span> at risk</div><div className="tally miss"><span className="n">{tally.miss}</span> missed</div></div>
        </div>
        <VolumesStrip reports={period.reports} prev={prev?.reports} />
        {(() => {
          const kpiScopes = activeScopes.filter((s) => s.criteria);
          const slaRows = kpiScopes.filter((s) => s.sla?.length).flatMap((s) => s.sla);
          const ownSla = activeScopes.filter((s) => s.kind === "tasks");
          const kpiCard = (sc) => (
            <div key={sc.key}>
              <div className="section-h"><h2>{sc.name}</h2><span className="rule" /></div>
              {!sc.criteria.some((crit) => period.scopes?.[sc.key]?.crit?.[crit.key]) && <EmptyScope scopeKey={sc.key} items={sc.criteria.map((c) => c.name)} kind="kpi" />}
              <div className="grid">
                {sc.criteria.map((crit) => {
                  const cell = period.scopes?.[sc.key]?.crit?.[crit.key]; if (!cell) return null;
                  const rule = effectiveRule(client, sc.key, crit); const st = critStatus(crit, cell, rule);
                  const isTxt = ["status", "yesno", "text"].includes(crit.type);
                  let delta = null; const pv = prev?.scopes?.[sc.key]?.crit?.[crit.key]?.value; if (isNum(pv) && isNum(cell.value)) delta = cell.value - pv;
                  return (
                    <div key={crit.key} className={"tile " + st}>
                      <div className="m-top"><span className="m-name">{crit.name}</span>{st !== "info" && <span className={"chip " + st}>{stLabel(st)}</span>}</div>
                      <div className={"m-val" + (isTxt ? " txt" : "")}>{critDisplay(crit, cell)}</div>
                      <div className="m-target">{isTxt ? (crit.std || "") : targetText(rule)}</div>
                      <div className="m-foot">
                        <span className="m-note">{cell.method === "upload" ? "from report" : "manual"}{cell.by ? ` · ${cell.by}` : ""}</span>
                        <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          {delta != null && delta !== 0 && <span className={"delta " + deltaCls(delta, rule)}>{delta > 0 ? "▲" : "▼"} {fmtVal(Math.abs(delta), rule.fmt)}</span>}
                          <button className="gear" title="How this is calculated" onClick={() => setInfo({ scopeKey: sc.key, crit, cell, rule })}>ⓘ</button>
                          {can("targets:edit") && !isTxt && <button className="gear" title="Edit target" onClick={() => setEditing({ scopeKey: sc.key, crit })}>⚙</button>}
                        </span>
                      </div>
                    </div>
                  );
                })}
              </div>
              {sc.hasAging && period.scopes?.[sc.key]?.aging?.rows?.some((r) => isNum(n(r.start)) || isNum(n(r.end)) || isNum(n(r.amount))) && <AgingPanel aging={period.scopes[sc.key].aging} prevAging={prev?.scopes?.[sc.key]?.aging} periodLabel={period.label} prevLabel={prev?.label} />}
            </div>
          );
          return <>
            {kpiScopes.length > 0 && <div className="group-h">KPI</div>}
            {kpiScopes.map(kpiCard)}
            {(period.reports?.payment?.payerPay || period.reports?.ar?.payerAR || period.reports?.denial?.payerDenialCnt) && <div><div className="section-h"><h2>Payer analysis (Pareto)</h2><span className="rule" /></div><ParetoSet reports={period.reports} scopes={activeScopes} /></div>}
            {(slaRows.length > 0 || ownSla.length > 0) && <div className="group-h">Service Level Agreement (SLA)</div>}
            {slaRows.length > 0 && <div><div className="section-h"><h2>Service Level Agreement</h2><span className="rule" /></div><TaskPanel rows={period.scopes?.sla?.tasks || []} scopeKey="sla" names={slaRows} /></div>}
            {activeScopes.some((s) => s.key === "ar") && period.reports?.ar?.aging && <div><div className="section-h"><h2>Total Inventory</h2><span className="rule" /></div><TotalInventory data={period.reports.ar} /></div>}
            {activeScopes.some((s) => s.key === "ar") && period.reports?.arprev?.aging && <div><div className="section-h"><h2>AR Inventory Aging Comparison</h2><span className="rule" /></div><ARAgingComparison prev={period.reports.arprev} curr={period.reports.ar} /></div>}
            {activeScopes.some((s) => s.key === "ar") && period.reports?.ar?.statusBreakdown && <div><div className="section-h"><h2>Status Wise Breakdown</h2><span className="rule" /></div><StatusBreakdown data={period.reports.ar} /></div>}
            {activeScopes.some((s) => s.key === "ar") && period.reports?.ar?.statusAging && <div><div className="section-h"><h2>Status Wise Breakdown, Aging Wise</h2><span className="rule" /></div><StatusAgingMatrix data={period.reports.ar} /></div>}
            {activeScopes.some((s) => s.key === "ar") && period.reports?.ar?.statusBreakdown && <div><div className="section-h"><h2>Category Wise Breakdown</h2><span className="rule" /></div><CategoryBreakdown data={period.reports.ar} mapping={client.statusCategories} canEdit={false} /></div>}
            {ownSla.map((sc) => <div key={sc.key}><div className="section-h"><h2>{sc.name}</h2><span className="rule" /></div><TaskPanel rows={period.scopes?.[sc.key]?.tasks || []} scopeKey={sc.key} names={sc.tasks} /></div>)}
          </>;
        })()}
        {can("period:delete") && <div style={{ marginTop: 22, display: "flex", justifyContent: "flex-end" }}><button className="btn ghost" onClick={() => setDelPeriod(true)}>Remove period</button></div>}
        {delPeriod && <Modal title="Remove period" onClose={() => setDelPeriod(false)}>
          <p style={{ fontSize: 13, color: "var(--ink-2)", marginTop: 0 }}>Remove <b>{client.periods?.[pIdx]?.label || "this period"}</b>? The uploaded reports and the values recorded against it go with it.</p>
          <div className="modal-actions"><button className="btn ghost" onClick={() => setDelPeriod(false)}>Cancel</button><button className="btn danger solid" onClick={async () => { const next = { ...client, periods: client.periods.filter((_, i) => i !== pIdx) }; setClient(next); setPIdx(0); setDelPeriod(false); await sSet(clientKey(client.id), next); }}>Remove period</button></div>
        </Modal>}
      </>}
      {info && <InfoModal data={info} onClose={() => setInfo(null)} />}
      {editing && can("targets:edit") && <EditRule scopeName={SCOPE_BY_KEY[editing.scopeKey].name} crit={editing.crit} rule={effectiveRule(client, editing.scopeKey, editing.crit)} onClose={() => setEditing(null)} onSave={async (nr) => { const next = { ...client, rules: { ...(client.rules || {}), [ruleKey(editing.scopeKey, editing.crit.key)]: nr } }; await sSet(clientKey(client.id), next); setClient(next); setEditing(null); }} />}
    </div>
  );
}
function StatusAgingMatrix({ data }) {
  const mat = data?.statusAging || {};
  const statuses = Object.keys(mat).sort((a, b) => { const an = a === "(blank)" ? 1 : 0, bn = b === "(blank)" ? 1 : 0; return an - bn || a.localeCompare(b); });
  const dollar = (v) => (v ? "$" + Math.round(v).toLocaleString() : "—");
  const rowTot = (st) => AR_BUCKETS.reduce((s, b) => s + (mat[st]?.[b] || 0), 0);
  const colTot = (b) => statuses.reduce((s, st) => s + (mat[st]?.[b] || 0), 0);
  const grand = statuses.reduce((s, st) => s + rowTot(st), 0);
  return (
    <div className="panel" style={{ marginTop: 6 }}>
      <div style={{ padding: "11px 14px", fontSize: 12.5, fontWeight: 600, borderBottom: "1px solid var(--line-2)" }}>Status Wise Breakdown, Aging Wise <span style={{ color: "var(--ink-3)", fontWeight: 400 }}>· amounts</span></div>
      <div style={{ overflowX: "auto" }}>
        <table className="cmptbl matx">
          <thead><tr><th>Status (Action)</th>{AR_BUCKETS.map((b) => <th key={b} style={{ textAlign: "right" }}>{b}</th>)}<th style={{ textAlign: "right" }}>Grand Total</th></tr></thead>
          <tbody>
            {statuses.map((st) => (<tr key={st}><td style={{ whiteSpace: "nowrap" }}>{st}</td>{AR_BUCKETS.map((b) => <td key={b} className="mono" style={{ textAlign: "right" }}>{dollar(mat[st]?.[b])}</td>)}<td className="mono" style={{ textAlign: "right", fontWeight: 600 }}>{dollar(rowTot(st))}</td></tr>))}
            <tr className="tot"><td>Grand Total</td>{AR_BUCKETS.map((b) => <td key={b} className="mono" style={{ textAlign: "right" }}>{dollar(colTot(b))}</td>)}<td className="mono" style={{ textAlign: "right" }}>{dollar(grand)}</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}
const CATEGORIES = ["Client", "Payer", "Purview"];
function guessCategory(status) {
  const s = String(status).toLowerCase();
  if (/client/.test(s) || /non.?workable/.test(s) || /patient\s*resp/.test(s)) return "Client";
  if (/portal/.test(s) || /registered/.test(s) || /\bpayer\b/.test(s) || /insurance/.test(s)) return "Payer";
  if (/rework|untouch|follow.?up|assistance|posting|submission|submit|adjust|flip|hold|process|call.?ins|denied|resubmit|appeal|sent to|review|rebill|correct/.test(s)) return "Purview";
  return null;
}
function CategoryBreakdown({ data, mapping, canEdit, onEdit }) {
  const sb = data?.statusBreakdown || {};
  const catOf = (st) => (mapping && mapping[st]) || guessCategory(st) || "Unassigned";
  const groups = { Client: [], Payer: [], Purview: [], Unassigned: [] };
  for (const [st, v] of Object.entries(sb)) groups[catOf(st)].push([st, v]);
  const dollar = (v) => "$" + Math.round(v).toLocaleString();
  const order = ["Client", "Payer", "Purview", "Unassigned"];
  let gCnt = 0, gAmt = 0;
  return (
    <div className="panel" style={{ marginTop: 6 }}>
      <div style={{ padding: "11px 14px", fontSize: 12.5, fontWeight: 600, borderBottom: "1px solid var(--line-2)", display: "flex", justifyContent: "space-between", alignItems: "center" }}><span>Category Wise Breakdown <span style={{ color: "var(--ink-3)", fontWeight: 400 }}>· Client / Payer / Purview</span></span>{canEdit && <button className="btn sm" onClick={onEdit}>Edit mapping</button>}</div>
      <table className="cmptbl"><thead><tr><th>Bucket</th><th>Status</th><th style={{ textAlign: "right" }}>Claim #</th><th style={{ textAlign: "right" }}>Total Balance</th></tr></thead>
        <tbody>
          {order.filter((cat) => groups[cat].length).map((cat) => {
            const rows = groups[cat].sort((a, b) => b[1].amount - a[1].amount);
            const cCnt = rows.reduce((s, [, v]) => s + v.count, 0), cAmt = rows.reduce((s, [, v]) => s + v.amount, 0);
            gCnt += cCnt; gAmt += cAmt;
            return rows.map(([st, v], i) => (
              <tr key={cat + st} className={cat === "Unassigned" ? "unassigned-row" : ""}>
                {i === 0 ? <td rowSpan={rows.length} style={{ fontWeight: 600, verticalAlign: "top", borderRight: "1px solid var(--line-2)", color: cat === "Unassigned" ? "var(--warn)" : "var(--ink)" }}>{cat}<div style={{ fontSize: 10.5, fontWeight: 400, color: "var(--ink-3)", marginTop: 2 }}>{cCnt} · {dollar(cAmt)}</div></td> : null}
                <td>{st}</td><td className="mono" style={{ textAlign: "right" }}>{v.count.toLocaleString()}</td><td className="mono" style={{ textAlign: "right" }}>{dollar(v.amount)}</td>
              </tr>
            ));
          })}
          <tr className="tot"><td colSpan={2}>Total</td><td className="mono" style={{ textAlign: "right" }}>{gCnt.toLocaleString()}</td><td className="mono" style={{ textAlign: "right" }}>{dollar(gAmt)}</td></tr>
        </tbody></table>
      {groups.Unassigned.length > 0 && canEdit && <div style={{ padding: "8px 14px", fontSize: 11, color: "var(--warn)" }}>{groups.Unassigned.length} status{groups.Unassigned.length > 1 ? "es" : ""} not yet categorized. Use “Edit mapping” to assign them.</div>}
    </div>
  );
}
function CategoryMappingEditor({ statuses, mapping, onClose, onSave }) {
  const [map, setMap] = useState(() => { const m = {}; statuses.forEach((s) => { m[s] = (mapping && mapping[s]) || guessCategory(s) || ""; }); return m; });
  return (<Modal onClose={onClose} title="Category mapping" desc="Assign each status (from the AR inventory Action column) to a category. Saved for this client and reused each period.">
    <div style={{ maxHeight: 340, overflowY: "auto", margin: "4px 0" }}>
      {statuses.map((s) => (
        <div key={s} style={{ display: "grid", gridTemplateColumns: "1fr 130px", gap: 10, alignItems: "center", padding: "5px 0", borderBottom: "1px solid var(--line-2)" }}>
          <span style={{ fontSize: 12.5 }}>{s}</span>
          <select value={map[s] || ""} onChange={(e) => setMap((m) => ({ ...m, [s]: e.target.value }))}>
            <option value="">— Unassigned —</option>
            {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
      ))}
    </div>
    <Actions onClose={onClose} onOk={() => onSave(map)} okLabel="Save mapping" />
  </Modal>);
}
function ParetoPanel({ title, data, fmt, explain }) {
  const [info, setInfo] = useState(false);
  const entries = Object.entries(data || {}).filter(([k, v]) => v > 0 && k !== "(blank)").sort((a, b) => b[1] - a[1]);
  if (!entries.length) return null;
  const total = entries.reduce((s, [, v]) => s + v, 0);
  const TOPN = 8;
  const top = entries.slice(0, TOPN);
  const rest = entries.slice(TOPN);
  if (rest.length) top.push(["Others", rest.reduce((s, [, v]) => s + v, 0)]);
  const max = Math.max(...top.map(([, v]) => v), 1);
  let cum = 0; const rows = top.map(([name, v]) => { cum += v; return { name, v, pct: v / total * 100, cum: cum / total * 100 }; });
  const fmtV = fmt === "money" ? (v) => "$" + Math.round(v).toLocaleString() : (v) => Math.round(v).toLocaleString();
  const fmtShort = fmt === "money" ? (v) => (v >= 1e6 ? "$" + (v / 1e6).toFixed(1) + "M" : v >= 1e3 ? "$" + (v / 1e3).toFixed(0) + "k" : "$" + Math.round(v)) : (v) => Math.round(v).toLocaleString();
  const short = (s) => (s.length > 12 ? s.slice(0, 11) + "…" : s);
  const W = 620, H = 252, padL = 8, padR = 40, padT = 34, padB = 58;
  const cw = W - padL - padR, ch = H - padT - padB;
  const step = cw / rows.length, bw = step * 0.56;
  const yBar = (v) => padT + ch - (v / max) * ch;
  const yPct = (p) => padT + ch - (p / 100) * ch;
  const cx = (i) => padL + i * step + step / 2;
  const line = rows.map((r, i) => `${cx(i)},${yPct(r.cum)}`).join(" ");
  const eighty = rows.findIndex((r) => r.cum >= 80);
  return (
    <div className="panel" style={{ marginTop: 6 }}>
      <div style={{ padding: "11px 14px", fontSize: 12.5, fontWeight: 600, borderBottom: "1px solid var(--line-2)", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span>{title}</span>
        <button className="infobtn" title="What this shows" onClick={() => setInfo(!info)} style={{ fontSize: 14 }}>ⓘ</button>
      </div>
      {info && <div style={{ padding: "10px 14px", fontSize: 12, color: "var(--ink-2)", background: "var(--steel-soft)", lineHeight: 1.5 }}>{explain}</div>}
      <div style={{ padding: "10px 8px 4px", overflowX: "auto" }}>
        <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", minWidth: 560, height: "auto" }} role="img">
          <line x1={padL} y1={yPct(80)} x2={padL + cw} y2={yPct(80)} stroke="var(--ink-3)" strokeWidth="1" strokeDasharray="4 3" />
          <text x={padL + cw + 3} y={yPct(80) + 3} fontSize="9.5" fill="var(--ink-3)">80%</text>
          {[0, 25, 50, 75, 100].map((p) => <text key={p} x={padL + cw + 3} y={yPct(p) + 3} fontSize="9" fill="var(--ink-3)">{p}%</text>)}
          {rows.map((r, i) => <rect key={i} x={cx(i) - bw / 2} y={yBar(r.v)} width={bw} height={padT + ch - yBar(r.v)} rx="2" fill="var(--steel)" opacity={i <= eighty || eighty < 0 ? 1 : 0.45} />)}
          {rows.map((r, i) => (<g key={"lab" + i}>
            <text x={cx(i)} y={yBar(r.v) - 14} fontSize="9" fontWeight="600" fill="var(--ink)" textAnchor="middle">{fmtShort(r.v)}</text>
            <text x={cx(i)} y={yBar(r.v) - 4} fontSize="8.5" fill="var(--steel)" textAnchor="middle">{r.pct.toFixed(1)}%</text>
          </g>))}
          <polyline points={line} fill="none" stroke="var(--navy)" strokeWidth="1.6" />
          {rows.map((r, i) => <circle key={i} cx={cx(i)} cy={yPct(r.cum)} r="2.6" fill="var(--navy)" />)}
          {rows.map((r, i) => <text key={i} x={cx(i)} y={H - padB + 14} fontSize="9.5" fill="var(--ink-2)" transform={`rotate(-32 ${cx(i)} ${H - padB + 14})`} textAnchor="end">{short(r.name)}</text>)}
        </svg>
      </div>
      <table className="cmptbl" style={{ marginTop: 2 }}><thead><tr><th>Payer</th><th style={{ textAlign: "right" }}>{fmt === "money" ? "Amount" : "Count"}</th><th style={{ textAlign: "right" }}>% of total</th><th style={{ textAlign: "right" }}>Cumulative %</th></tr></thead>
        <tbody>{rows.map((r, i) => (<tr key={i}><td>{r.name}</td><td className="mono" style={{ textAlign: "right" }}>{fmtV(r.v)}</td><td className="mono" style={{ textAlign: "right" }}>{r.pct.toFixed(1)}%</td><td className="mono" style={{ textAlign: "right", color: r.cum >= 80 ? "var(--ink-3)" : "var(--ink)" }}>{r.cum.toFixed(1)}%</td></tr>))}</tbody></table>
    </div>
  );
}
function ParetoSet({ reports, scopes }) {
  const has = (k) => scopes.some((s) => s.key === k);
  const R = reports || {};
  return (<>
    {has("pp") && R.payment?.payerPay && <ParetoPanel title="Top Payers by Payment Received" data={R.payment.payerPay} fmt="money" explain="Bars show how much each payer paid, largest first. The line is the running share of total payments. It tells you which few payers make up most of your money, so a small number of payers usually drive the bulk of collections." />}
    {has("ar") && R.ar?.payerAR && <ParetoPanel title="Top Payers by AR Outstanding" data={R.ar.payerAR} fmt="money" explain="Bars show the outstanding AR balance per payer, largest first. The line is the running share of total AR. It points to where the most money is stuck, so your team can focus collection effort on the payers that carry most of the outstanding balance." />}
    {has("ar") && R.denial?.payerDenialCnt && <ParetoPanel title="Top Payers by Denials (count)" data={R.denial.payerDenialCnt} fmt="count" explain="Bars show how many denials came from each payer, largest first. The line is the running share of total denials. It shows which payers deny the most, so you know where to push back or fix root causes first." />}
  </>);
}
function StatusBreakdown({ data }) {
  const sb = data?.statusBreakdown || {};
  const rows = Object.entries(sb).sort((a, b) => { const an = a[0] === "(blank)" ? 1 : 0, bn = b[0] === "(blank)" ? 1 : 0; return an - bn || a[0].localeCompare(b[0]); });
  const sumCnt = rows.reduce((s, [, v]) => s + v.count, 0), sumAmt = rows.reduce((s, [, v]) => s + v.amount, 0);
  const dollar = (v) => "$" + Math.round(v).toLocaleString();
  return (
    <div className="panel" style={{ marginTop: 6 }}>
      <div style={{ padding: "11px 14px", fontSize: 12.5, fontWeight: 600, borderBottom: "1px solid var(--line-2)" }}>Status Wise Breakdown <span style={{ color: "var(--ink-3)", fontWeight: 400 }}>· by Action</span></div>
      <table className="cmptbl"><thead><tr><th>Status (Action)</th><th style={{ textAlign: "right" }}>Claim #</th><th style={{ textAlign: "right" }}>Total Amount</th></tr></thead>
        <tbody>
          {rows.map(([st, v]) => (<tr key={st}><td>{st}</td><td className="mono" style={{ textAlign: "right" }}>{v.count.toLocaleString()}</td><td className="mono" style={{ textAlign: "right" }}>{dollar(v.amount)}</td></tr>))}
          <tr className="tot"><td>Grand Total</td><td className="mono" style={{ textAlign: "right" }}>{sumCnt.toLocaleString()}</td><td className="mono" style={{ textAlign: "right" }}>{dollar(sumAmt)}</td></tr>
        </tbody></table>
    </div>
  );
}
function TotalInventory({ data }) {
  const norm = (obj) => { const g = {}; Object.entries(obj || {}).forEach(([b, v]) => { const k = canonBucket(b); g[k] = (g[k] || 0) + v; }); return g; };
  const amt = norm(data?.aging), cnt = norm(data?.agingCount);
  const sumAmt = Object.values(amt).reduce((a, b) => a + b, 0), sumCnt = Object.values(cnt).reduce((a, b) => a + b, 0);
  const dollar = (v) => "$" + Math.round(v).toLocaleString();
  return (
    <div className="panel" style={{ marginTop: 6 }}>
      <div style={{ padding: "11px 14px", fontSize: 12.5, fontWeight: 600, borderBottom: "1px solid var(--line-2)" }}>Total Inventory, current period</div>
      <table className="cmptbl"><thead><tr><th>Aging Bucket</th><th style={{ textAlign: "right" }}>Claim #</th><th style={{ textAlign: "right" }}>Total Amount</th></tr></thead>
        <tbody>
          {AR_BUCKETS.map((b) => (<tr key={b}><td>{b}</td><td className="mono" style={{ textAlign: "right" }}>{cnt[b] || 0}</td><td className="mono" style={{ textAlign: "right" }}>{amt[b] ? dollar(amt[b]) : "—"}</td></tr>))}
          <tr className="tot"><td>Grand Total</td><td className="mono" style={{ textAlign: "right" }}>{sumCnt.toLocaleString()}</td><td className="mono" style={{ textAlign: "right" }}>{dollar(sumAmt)}</td></tr>
        </tbody></table>
    </div>
  );
}
function ARAgingComparison({ prev, curr }) {
  const norm = (aging) => { const g = {}; Object.entries(aging || {}).forEach(([b, v]) => { const k = canonBucket(b); g[k] = (g[k] || 0) + v; }); return g; };
  const pv = norm(prev?.aging), cv = norm(curr?.aging);
  const sumP = Object.values(pv).reduce((a, b) => a + b, 0), sumC = Object.values(cv).reduce((a, b) => a + b, 0);
  const comp = (p, c) => (p > 0 ? (p - c) / p * 100 : null);
  const dollar = (v) => "$" + Math.round(v).toLocaleString();
  const pct = (x) => (x == null ? "—" : x.toFixed(1) + "%");
  const cls = (x) => (x == null ? "" : x >= 0 ? "ok" : "miss");
  return (
    <div className="panel" style={{ marginTop: 6 }}>
      <div style={{ padding: "11px 14px", fontSize: 12.5, fontWeight: 600, borderBottom: "1px solid var(--line-2)" }}>AR Inventory Aging Comparison: Previous vs Current</div>
      <table className="cmptbl"><thead><tr><th>Aging Bucket</th><th style={{ textAlign: "right" }}>Previous</th><th style={{ textAlign: "right" }}>Current</th><th style={{ textAlign: "right" }}>Comparison</th></tr></thead>
        <tbody>
          {AR_BUCKETS.map((b) => { const p = pv[b] || 0, c = cv[b] || 0, x = comp(p, c); return (<tr key={b}><td>{b}</td><td className="mono" style={{ textAlign: "right" }}>{p ? dollar(p) : "—"}</td><td className="mono" style={{ textAlign: "right" }}>{c ? dollar(c) : "—"}</td><td className="mono" style={{ textAlign: "right" }}><span className={"cmp " + cls(x)}>{pct(x)}</span></td></tr>); })}
          <tr className="tot"><td>Grand Total</td><td className="mono" style={{ textAlign: "right" }}>{dollar(sumP)}</td><td className="mono" style={{ textAlign: "right" }}>{dollar(sumC)}</td><td className="mono" style={{ textAlign: "right" }}><span className={"cmp " + cls(comp(sumP, sumC))}>{pct(comp(sumP, sumC))}</span></td></tr>
          <tr className="cnt"><td>Count (claims)</td><td className="mono" style={{ textAlign: "right" }}>{prev?.count ?? "—"}</td><td className="mono" style={{ textAlign: "right" }}>{curr?.count ?? "—"}</td><td /></tr>
        </tbody></table>
    </div>
  );
}
function AgingPanel({ aging, prevAging, periodLabel, prevLabel }) {
  const val = (r) => (n(r.amount) != null ? n(r.amount) : n(r.end) != null ? n(r.end) : n(r.start) || 0);
  const curr = {}; (aging.rows || []).forEach((r) => { const v = val(r); if (isNum(v)) curr[r.bucket] = v; });
  const prev = {}; (prevAging?.rows || []).forEach((r) => { const v = val(r); if (isNum(v)) prev[r.bucket] = v; });
  const hasPrev = Object.keys(prev).length > 0;
  const lb = (b) => { const m = String(b).match(/\d+/); return m ? parseInt(m[0]) : 1e9; };
  const buckets = Array.from(new Set([...Object.keys(curr), ...Object.keys(prev)])).sort((a, b) => lb(a) - lb(b));
  const max = Math.max(...buckets.map((b) => Math.max(curr[b] || 0, prev[b] || 0)), 1);
  const cols = hasPrev ? "116px 1fr 84px 92px" : "130px 1fr 100px";
  return (
    <div className="panel" style={{ marginTop: 6 }}>
      <div style={{ padding: "12px 16px", fontSize: 12.5, fontWeight: 600, color: "var(--ink-2)", borderBottom: "1px solid var(--line-2)" }}>AR aging{hasPrev ? `: ${prevLabel || "previous"} vs ${periodLabel || "current"}` : " distribution"}</div>
      <div style={{ padding: "10px 16px" }}>
        {buckets.map((b, i) => {
          const cv = curr[b] || 0, pv = prev[b] || 0, d = cv - pv, bad = is90plus(b);
          return (
            <div key={i} style={{ display: "grid", gridTemplateColumns: cols, gap: 10, alignItems: "center", padding: "5px 0", fontSize: 12.5 }}>
              <span style={{ color: bad ? "var(--miss)" : "var(--ink-2)" }}>{b}</span>
              <span style={{ height: 14, background: "var(--line-2)", borderRadius: 3, overflow: "hidden", display: "flex" }}><span style={{ width: `${(cv / max) * 100}%`, background: bad ? "var(--miss)" : "var(--steel)" }} /></span>
              {hasPrev && <span className="mono" style={{ textAlign: "right", color: "var(--ink-3)" }}>${Math.round(pv).toLocaleString()}</span>}
              <span className="mono" style={{ textAlign: "right", color: "var(--ink-2)" }}>${Math.round(cv).toLocaleString()}{hasPrev && d !== 0 && <span style={{ color: "var(--ink-3)", fontSize: 10 }}> {d > 0 ? "+" : ""}{Math.round(d).toLocaleString()}</span>}</span>
            </div>
          );
        })}
        {hasPrev && <div style={{ display: "grid", gridTemplateColumns: cols, gap: 10, fontSize: 10, color: "var(--ink-3)", marginTop: 6, paddingTop: 6, borderTop: "1px solid var(--line-2)" }}><span /><span /><span style={{ textAlign: "right" }}>{prevLabel || "prev"}</span><span style={{ textAlign: "right" }}>{periodLabel || "current"}</span></div>}
      </div>
    </div>
  );
}
function VolumesStrip({ reports, prev }) {
  const Rr = reports || {}, Pp = prev || {};
  const money = (v) => "$" + Math.round(v).toLocaleString();
  const num = (v) => Math.round(v).toLocaleString();
  const items = [];
  const add = (label, val, pval, fmt) => { if (!isNum(val)) return; items.push({ label, val: fmt(val), delta: isNum(pval) ? val - pval : null, fmt }); };
  add("Payments received", Rr.payment?.payments, Pp.payment?.payments, money);
  add("Total AR", Rr.ar?.totalAR, Pp.ar?.totalAR, money);
  add("Charges", Rr.ar?.totalCharges, Pp.ar?.totalCharges, money);
  add("Denials", Rr.denial?.count, Pp.denial?.count, num);
  add("Posted claims", Rr.payment?.rows, Pp.payment?.rows, num);
  add("Submissions", Rr.submission?.count, Pp.submission?.count, num);
  if (!items.length) return null;
  return (<><div className="section-h"><h2>Source volumes</h2><span className="rule" /></div>
    <div className="volgrid">{items.map((it, i) => (
      <div key={i} className="volcard"><div className="vl">{it.label}</div><div className="vv">{it.val}</div>
        {it.delta != null && it.delta !== 0 && <div className="vd">{it.delta > 0 ? "▲" : "▼"} {it.fmt(Math.abs(it.delta))} vs prev</div>}</div>
    ))}</div></>);
}
function EmptyScope({ scopeKey, items, kind }) {
  const col = scopeColor(scopeKey);
  return (
    <div style={{ position: "relative", overflow: "hidden", borderRadius: 18, padding: "17px 20px", border: `1px solid ${col}3d`, boxShadow: "var(--sh-1)", background: `radial-gradient(120% 150% at 100% 0%, ${col}18, transparent 55%), linear-gradient(140deg,#fff 54%, ${col}10)` }}>
      <div style={{ display: "flex", gap: 12, alignItems: "center", marginBottom: items?.length ? 12 : 0 }}>
        <span style={{ width: 38, height: 38, flex: "none", borderRadius: 11, background: col + "26", color: col, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 10.5 }}>{scopeBadge(scopeKey)}</span>
        <div>
          <div style={{ fontSize: 13, fontWeight: 600 }}>No {kind === "kpi" ? "KPI values" : "SLA entries"} recorded for this period yet</div>
          <div style={{ fontSize: 11.5, color: "var(--ink-3)", marginTop: 2 }}>Open this client under <b>Clients</b> to enter or upload the numbers.</div>
        </div>
      </div>
      {items?.length ? <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
        {items.map((nm, i) => <span key={i} style={{ fontSize: 11, padding: "4px 11px", borderRadius: 999, background: "rgba(22,35,59,.035)", color: "var(--ink-3)", border: "1px dashed var(--line)" }}>{nm}</span>)}
      </div> : null}
    </div>
  );
}
function TaskPanel({ rows, scopeKey = "sla", names }) {
  const has = rows.some((r) => r.target || r.received || r.achieved || r.comment);
  if (!has) return <EmptyScope scopeKey={scopeKey} items={names || rows.map((r) => r.name)} kind="sla" />;
  return (<div className="panel"><table className="tasktbl"><thead><tr><th>Metric</th><th>Target</th><th>Received</th><th>Achieved</th><th>Achievement</th><th>Comment</th></tr></thead>
    <tbody>{rows.filter((r) => r.target || r.received || r.achieved || r.comment).map((r, i) => {
      const rec = n(r.received), ach = n(r.achieved); const pct = rec && ach != null ? ach / rec : null;
      const st = pct == null ? "info" : pct >= 0.9 ? "ok" : pct >= 0.7 ? "warn" : "miss";
      return (<tr key={i}><td style={{ color: "var(--ink)", fontWeight: 500 }}>{r.name}</td><td className="mono">{r.target || "—"}</td><td className="mono">{r.received || "—"}</td><td className="mono">{r.achieved || "—"}</td>
        <td>{pct != null ? <span className={"chip " + st}>{Math.round(pct * 100)}%</span> : "—"}</td><td style={{ fontSize: 12 }}>{r.comment || "—"}</td></tr>);
    })}</tbody></table></div>);
}

/* ============================ data entry =============================== */
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const MSHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
/* ============================ periods: months, weeks and days ============================
   Every upload belongs to a period: a month, a week (Monday to Sunday, and it may run into the next
   month), a single day, or any other dates. Periods are listed newest first, and each one is compared
   with the one before it of the same kind: this month with last month, this week with last week. */
const WEEK_START = 1; // Monday
const pad2 = (x) => String(x).padStart(2, "0");
const ymd = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const parseYmd = (s) => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || "")); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
const addDays = (d, k) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + k);
const weekStartOf = (d) => addDays(d, -((d.getDay() - WEEK_START + 7) % 7));
const daysBetween = (a, b) => Math.round((b - a) / 86400000);
function periodType(p) {
  if (p && (p.kind === "month" || p.kind === "week" || p.kind === "day")) return p.kind;
  const f = parseYmd(p?.range?.from), t = parseYmd(p?.range?.to); if (!f || !t) return "custom";
  if (+f === +t) return "day";
  if (f.getDate() === 1 && t.getFullYear() === f.getFullYear() && t.getMonth() === f.getMonth() && t.getDate() === new Date(t.getFullYear(), t.getMonth() + 1, 0).getDate()) return "month";
  if (daysBetween(f, t) === 6 && f.getDay() === WEEK_START) return "week";
  return "custom";
}
const TYPE_NAME = { month: "Monthly", week: "Weekly", day: "Daily", custom: "Other dates" };
const TYPE_UNIT = { month: "month", week: "week", day: "day", custom: "period" };
function rangeFor(kind, d) { // the month, week or day that holds date d
  if (kind === "month") return { from: ymd(new Date(d.getFullYear(), d.getMonth(), 1)), to: ymd(new Date(d.getFullYear(), d.getMonth() + 1, 0)) };
  if (kind === "week") { const st = weekStartOf(d); return { from: ymd(st), to: ymd(addDays(st, 6)) }; }
  return { from: ymd(d), to: ymd(d) };
}
function stepRange(kind, from, k) {
  const f = parseYmd(from);
  if (kind === "month") return rangeFor("month", new Date(f.getFullYear(), f.getMonth() + k, 1));
  if (kind === "week") return rangeFor("week", addDays(f, 7 * k));
  return rangeFor("day", addDays(f, k));
}
function stepsBetween(kind, fromA, fromB) { // how many months, weeks or days from A to B
  const a = parseYmd(fromA), b = parseYmd(fromB);
  if (kind === "month") return (b.getFullYear() - a.getFullYear()) * 12 + b.getMonth() - a.getMonth();
  return Math.round(daysBetween(a, b) / (kind === "week" ? 7 : 1));
}
function periodLabel(kind, from, to) {
  const f = parseYmd(from), t = parseYmd(to); if (!f || !t) return "";
  if (kind === "month") return `${MONTHS[f.getMonth()]} ${f.getFullYear()}`;
  if (kind === "week") return `Week of ${f.getDate()} ${MSHORT[f.getMonth()]} ${f.getFullYear()}`;
  if (kind === "day") return `${f.getDate()} ${MSHORT[f.getMonth()]} ${f.getFullYear()}`;
  return `${f.getDate()} ${MSHORT[f.getMonth()]}${f.getFullYear() !== t.getFullYear() ? " " + f.getFullYear() : ""} to ${t.getDate()} ${MSHORT[t.getMonth()]} ${t.getFullYear()}`;
}
function shortLabel(kind, from, to) { // for table headings
  const f = parseYmd(from), t = parseYmd(to); if (!f) return "";
  if (kind === "month") return `${MSHORT[f.getMonth()]} ${f.getFullYear()}`;
  if (kind === "week" || kind === "day") return `${f.getDate()} ${MSHORT[f.getMonth()]}`;
  return `${f.getDate()} ${MSHORT[f.getMonth()]} to ${t ? t.getDate() + " " + MSHORT[t.getMonth()] : ""}`;
}
const periodSort = (a, b) => String(b?.range?.to || "").localeCompare(String(a?.range?.to || "")) || String(b?.range?.from || "").localeCompare(String(a?.range?.from || "")) || String(b?.createdAt || "").localeCompare(String(a?.createdAt || ""));
const sortPeriods = (list) => [...(list || [])].sort(periodSort);
// the period before this one of the same kind: last month for a month, last week for a week
function prevSameType(periods, idx, k = 1) {
  const p = periods?.[idx]; if (!p) return null;
  const t = periodType(p); let seen = 0;
  for (let j = idx + 1; j < periods.length; j++) {
    const q = periods[j]; if (periodType(q) !== t) continue;
    if (p.range?.from && q.range?.to && q.range.to >= p.range.from) continue;
    if (++seen === k) return q;
  }
  return null;
}
const periodMarks = (periods) => new Set((periods || []).map((p) => periodType(p) + "|" + p.range?.from + "|" + p.range?.to));
const periodValue = (p) => (p ? { kind: periodType(p), from: p.range?.from, to: p.range?.to, label: p.label } : null);

// The short copy of a period used by the Overview and History: values, reasons and SLA rows, no report detail.
function periodLite(p) {
  const scopes = {};
  for (const [k, v] of Object.entries(p?.scopes || {})) {
    if (!v) continue; const o = {};
    if (v.crit) {
      o.crit = {};
      for (const [ck, c] of Object.entries(v.crit)) {
        if (!c) continue; const x = {};
        if (c.value !== undefined && c.value !== null && c.value !== "") x.value = c.value;
        if (c.method) x.method = c.method;
        if (c.reason) x.reason = String(c.reason).slice(0, 200);
        if (Object.keys(x).length) o.crit[ck] = x;
      }
    }
    if (Array.isArray(v.tasks)) o.tasks = v.tasks.map((t) => ({ name: t.name, target: t.target, received: t.received, achieved: t.achieved, comment: t.comment ? String(t.comment).slice(0, 200) : "" }));
    if (v.aging) o.aging = v.aging;
    scopes[k] = o;
  }
  const reports = {}; for (const [k, r] of Object.entries(p?.reports || {})) reports[k] = { stub: true, fileName: (r && r.fileName) || "" };
  return { id: p.id, kind: periodType(p), label: p.label, range: p.range, numDays: p.numDays, createdAt: p.createdAt, createdBy: p.createdBy, scopes, reports, _stub: true };
}
function periodSummary(client, p) {
  const sc = clientScorecard(client, p, null);
  return { overall: sc.pct, kpi: sc.kpiPct, sla: sc.slaPct, green: sc.counts.ok, amber: sc.counts.warn, red: sc.counts.miss };
}
const STATUS_WORD = { great: "Green", ok: "Green", warn: "Amber", miss: "Red" };
const round2 = (v) => Math.round(v * 100) / 100;
const inputsText = (inputs) => Object.entries(inputs || {}).map(([k, v]) => k + ": " + (typeof v === "number" ? v.toLocaleString() : v)).join("; ");
function slaHistoryRow(scopeName, r) {
  const rec = n(r.received), ach = n(r.achieved); if (!(isNum(rec) && rec > 0 && isNum(ach))) return null;
  const pc = ach / rec, st = pc >= 0.9 ? "ok" : pc >= 0.7 ? "warn" : "miss";
  return { scope: scopeName, metric: r.name, value: round2(pc * 100), unit: "%", shown: Math.round(pc * 100) + "%", target: "≥ 90%", status: STATUS_WORD[st], source: "Typed",
    inputs: "received: " + rec.toLocaleString() + "; achieved: " + ach.toLocaleString() + (r.target ? "; target: " + r.target : ""), reason: r.comment || "" };
}
// One row per KPI or SLA line in a period. The same rows go to the KPI History tab and the Excel download.
function historyRows(client, p) {
  const rows = [];
  const active = (client?.scopes || []).map((k) => SCOPE_BY_KEY[k]).filter(Boolean);
  for (const sc of active) {
    if (sc.criteria) for (const crit of sc.criteria) {
      const cell = p?.scopes?.[sc.key]?.crit?.[crit.key]; if (!cell || cell.value == null || cell.value === "") continue;
      const rule = effectiveRule(client, sc.key, crit), st = critStatus(crit, cell, rule);
      const isTxt = ["status", "yesno", "text"].includes(crit.type), fmt = critFmt(crit);
      let value = null, unit = "";
      if (!isTxt && isNum(cell.value)) { if (fmt === FMT.PERCENT) { value = round2(cell.value * 100); unit = "%"; } else if (fmt === FMT.DAYS) { value = round2(cell.value); unit = "days"; } else value = round2(cell.value); }
      rows.push({ scope: sc.name, metric: crit.name, value, unit, shown: critDisplay(crit, cell), target: isTxt ? (crit.std || "") : targetText(rule), status: STATUS_WORD[st] || "",
        source: cell.method === "upload" ? "Uploaded" : "Typed", inputs: inputsText(cell.inputs), reason: cell.reason || "" });
    }
    if (sc.kind === "tasks") for (const r of p?.scopes?.[sc.key]?.tasks || []) { const h = slaHistoryRow(sc.name, r); if (h) rows.push(h); }
  }
  for (const r of p?.scopes?.sla?.tasks || []) { const h = slaHistoryRow("Service Level Agreement", r); if (h) rows.push(h); }
  return rows;
}
const FIGURE_LABELS = {
  aging: "AR by aging bucket", agingCount: "Claims by aging bucket", statusBreakdown: "AR by status", statusAging: "AR by status and aging", payerAR: "AR by payer",
  totalAR: "Total AR", totalCharges: "Total charges", count: "Claims counted", claims: "Claims", appeals: "Appeals", byDept: "By department", rows: "Rows in the file",
  callsTotal: "Calls in total", callsPicked: "Calls picked up", callsMissed: "Calls missed", callsDropped: "Calls dropped", callsOther: "Other calls",
  cbDone: "Callbacks made", cbTotal: "Missed calls to call back", corrDone: "Correspondence sent", corrTotal: "Correspondence received", faxDone: "Fax and email completed", faxTotal: "Fax and email received",
  evCompleted: "Verifications completed", evTotal: "Verification requests", evUrgent: "Urgent requests", evRegular: "Regular requests",
  insurancePaid: "Insurance payments", patientPaid: "Patient payments", ledgerAllowed: "Allowed amount", ledgerUnder: "Underpaid amount", ledgerVariance: "Variance amount", varianceAmt: "Variance amount",
  payerDenialAmt: "Denied amount by payer", payerDenialCnt: "Denials by payer", payerPay: "Payments by payer", postingIn: "Amount received for posting", recoupAmt: "Amount recouped", writeoff: "Write-off",
};
const humanKey = (k) => FIGURE_LABELS[k] || String(k).replace(/([a-z])([A-Z])/g, "$1 $2").replace(/^./, (m) => m.toUpperCase());
const SKIP_FIGURES = new Set(["fileName", "raw", "rows_", "columns", "missing", "ambiguous", "dateIdx", "cols", "agingSource", "stub"]);
// Every total and breakdown worked out from the period's reports, for the Report Details tab
function breakdownRows(p) {
  const out = [];
  for (const [rk, rep] of Object.entries(p?.reports || {})) {
    if (!rep || typeof rep !== "object" || rep.stub) continue;
    const report = REPORT_BY_KEY[rk]?.name || rk;
    if (rep.fileName) out.push({ report, figure: "File used", item: String(rep.fileName), value: null });
    for (const [k, v] of Object.entries(rep)) {
      if (SKIP_FIGURES.has(k)) continue;
      if (isNum(v)) { out.push({ report, figure: humanKey(k), item: "", value: v }); continue; }
      if (!v || typeof v !== "object" || Array.isArray(v)) continue;
      let c = 0;
      for (const [ik, iv] of Object.entries(v)) {
        if (c >= 120) break;
        if (isNum(iv)) { out.push({ report, figure: humanKey(k), item: ik, value: iv }); c++; }
        else if (iv && typeof iv === "object" && !Array.isArray(iv)) for (const [jk, jv] of Object.entries(iv)) if (isNum(jv) && c < 120) { out.push({ report, figure: humanKey(k), item: ik + " · " + jk, value: jv }); c++; }
      }
    }
  }
  for (const r of p?.scopes?.ar?.aging?.rows || []) for (const f of ["amount", "start", "end"]) if (isNum(n(r[f]))) out.push({ report: "AR aging as recorded", figure: f === "amount" ? "Amount" : f === "start" ? "Start of period" : "End of period", item: r.bucket || "", value: n(r[f]) });
  return out.slice(0, 2500);
}

/* The period picker: Month, Week (Monday to Sunday, may cross into the next month), Day or other dates.
   Months, weeks and days that already have data carry a small dot. */
const PICK_TABS = [["month", "Month"], ["week", "Week"], ["day", "Day"], ["custom", "Other dates"]];
function PeriodPicker({ value, onSelect, types = ["month", "week", "day", "custom"], marks, placeholder = "Pick a period", align = "left" }) {
  const today = new Date();
  const [open, setOpen] = useState(false);
  const [tab, setTab] = useState(types[0]);
  const [vy, setVy] = useState(today.getFullYear());
  const [vm, setVm] = useState(today.getMonth());
  const [cs, setCs] = useState(null); const [ce, setCe] = useState(null);
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return;
    setTab(types.includes(value?.kind) ? value.kind : types[0]);
    const b = parseYmd(value?.from) || today; setVy(b.getFullYear()); setVm(b.getMonth());
    setCs(value?.kind === "custom" ? parseYmd(value.from) : null); setCe(value?.kind === "custom" ? parseYmd(value.to) : null);
    const onDown = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    const onKey = (e) => { if (e.key === "Escape") setOpen(false); };
    document.addEventListener("mousedown", onDown); document.addEventListener("keydown", onKey);
    return () => { document.removeEventListener("mousedown", onDown); document.removeEventListener("keydown", onKey); };
  }, [open]);
  const has = (kind, from, to) => !!marks && marks.has(kind + "|" + from + "|" + to);
  const pick = (kind, from, to) => { onSelect({ kind, from, to, label: periodLabel(kind, from, to), numDays: daysBetween(parseYmd(from), parseYmd(to)) + 1 }); setOpen(false); };
  const navMonth = (k) => { const d = new Date(vy, vm + k, 1); setVy(d.getFullYear()); setVm(d.getMonth()); };
  const lastDay = new Date(vy, vm + 1, 0), weeks = [];
  for (let st = weekStartOf(new Date(vy, vm, 1)); st <= lastDay; st = addDays(st, 7)) weeks.push(Array.from({ length: 7 }, (_, i) => addDays(st, i)));
  const clickCustom = (d) => { if (!cs || ce) { setCs(d); setCe(null); } else if (d < cs) { setCe(cs); setCs(d); } else setCe(d); };
  const inCustom = (d) => cs && (ce ? d >= cs && d <= ce : +d === +cs);
  return (
    <div ref={ref} className="pp">
      <button className={"btn pp-btn" + (open ? " on" : "")} onClick={() => setOpen(!open)} aria-haspopup="dialog" aria-expanded={open}>
        <span>{value?.label || placeholder}</span>
        <svg width="11" height="11" viewBox="0 0 12 12" aria-hidden="true"><path d="M2.5 4.5 6 8l3.5-3.5" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" /></svg>
      </button>
      {open && <div className={"calpop pp-pop" + (align === "right" ? " right" : "")} role="dialog" aria-label="Pick a period">
        {types.length > 1 && <div className="seg pp-tabs">{PICK_TABS.filter(([k]) => types.includes(k)).map(([k, l]) => <button key={k} className={"segbtn" + (tab === k ? " on" : "")} onClick={() => setTab(k)}>{l}</button>)}</div>}
        {tab === "month" ? <>
          <div className="calhead"><button className="calnav" aria-label="Previous year" onClick={() => setVy(vy - 1)}>‹</button><span>{vy}</span><button className="calnav" aria-label="Next year" onClick={() => setVy(vy + 1)}>›</button></div>
          <div className="pp-months">{MSHORT.map((m, i) => {
            const r = rangeFor("month", new Date(vy, i, 1)); const on = value?.kind === "month" && value.from === r.from;
            return <button key={m} className={"pp-month" + (on ? " on" : "")} onClick={() => pick("month", r.from, r.to)}>{m}{has("month", r.from, r.to) && <i className="pp-dot" title="Has data" />}</button>;
          })}</div>
        </> : <>
          <div className="calhead"><button className="calnav" aria-label="Previous month" onClick={() => navMonth(-1)}>‹</button><span>{MONTHS[vm]} {vy}</span><button className="calnav" aria-label="Next month" onClick={() => navMonth(1)}>›</button></div>
          <div className="calgrid caldow">{["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"].map((d) => <span key={d} className="dow">{d}</span>)}</div>
          <div className="pp-weeks">{weeks.map((wk) => {
            const wr = rangeFor("week", wk[0]), weekOn = tab === "week" && value?.kind === "week" && value.from === wr.from;
            return (
              <div key={wr.from} className={"pp-week" + (tab === "week" ? " pick" : "") + (weekOn ? " on" : "")} onClick={tab === "week" ? () => pick("week", wr.from, wr.to) : undefined}>
                {wk.map((d) => {
                  const k = ymd(d), out = d.getMonth() !== vm;
                  const dayOn = (tab === "day" && value?.kind === "day" && value.from === k) || (tab === "custom" && inCustom(d));
                  const edge = tab === "custom" && ((cs && +d === +cs) || (ce && +d === +ce));
                  return <button key={k} className={"calday" + (out ? " out" : "") + (dayOn ? " in" : "") + (edge || (dayOn && tab === "day") ? " edge" : "")}
                    onClick={tab === "day" ? () => pick("day", k, k) : tab === "custom" ? () => clickCustom(d) : undefined}>{d.getDate()}{tab === "day" && has("day", k, k) && <i className="pp-dot" />}</button>;
                })}
                {tab === "week" && has("week", wr.from, wr.to) && <i className="pp-dot wk" title="Has data" />}
              </div>
            );
          })}</div>
          {tab === "week" && <div className="pp-hint">Weeks run Monday to Sunday. Click any day to pick its week.</div>}
          {tab === "custom" && <div className="calfoot">
            <span className="pp-hint" style={{ margin: 0 }}>{cs ? (ce ? periodLabel("custom", ymd(cs), ymd(ce)) : "Now click the last day") : "Click the first day"}</span>
            <button className="btn primary sm" disabled={!cs} onClick={() => pick("custom", ymd(cs), ymd(ce || cs))}>Use these dates</button>
          </div>}
        </>}
      </div>}
    </div>
  );
}
/* -------- choosing a period inside a client: the picker, the four newest, and the rest in a list -------- */
function PeriodSelect({ periods, pIdx, onIndex, emptyLabel }) {
  return (
    <select className="pp-all" value={emptyLabel && pIdx < 4 ? "" : String(pIdx)} onChange={(e) => { if (e.target.value !== "") onIndex(Number(e.target.value)); }} aria-label="Saved periods">
      {emptyLabel && <option value="">{emptyLabel}</option>}
      {["month", "week", "day", "custom"].map((t) => {
        const list = periods.map((p, i) => [p, i]).filter(([p]) => periodType(p) === t);
        return list.length ? <optgroup key={t} label={TYPE_NAME[t]}>{list.map(([p, i]) => <option key={p.id} value={i}>{p.label}</option>)}</optgroup> : null;
      })}
    </select>
  );
}
function PeriodBar({ client, pIdx, period, onPick, onIndex, numDays, onDays }) {
  const periods = client?.periods || [];
  const marks = useMemo(() => periodMarks(periods), [periods]);
  return (
    <div className="card pp-card" style={{ marginBottom: 16 }}>
      <div className="row" style={{ justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
        <div className="row" style={{ gap: 10, flexWrap: "wrap" }}>
          <span className="pp-lab">Period</span>
          <PeriodPicker value={periodValue(period)} onSelect={onPick} marks={marks} placeholder="Pick a month, week or day" />
          {period && <span className="pp-type">{TYPE_NAME[periodType(period)]}</span>}
        </div>
        {period && onDays && <div className="row"><span style={{ fontSize: 12, color: "var(--ink-2)" }}>Days</span><input value={numDays} onChange={(e) => onDays(n(e.target.value) || "")} className="mono" style={{ width: 58 }} /></div>}
      </div>
      {periods.length > 0 && <div className="pp-recent">
        <span className="pp-lab">Saved</span>
        {periods.slice(0, 4).map((p, i) => <button key={p.id} className={"pp-chip" + (i === pIdx ? " on" : "")} onClick={() => onIndex(i)}>{p.label}</button>)}
        {periods.length > 4 && <PeriodSelect periods={periods} pIdx={pIdx} onIndex={onIndex} emptyLabel={`All ${periods.length} periods`} />}
      </div>}
    </div>
  );
}
function ViewSwitch({ tab, onTab }) {
  return (<div className="seg view-switch" role="tablist">
    <button role="tab" aria-selected={tab === "period"} className={"segbtn" + (tab === "period" ? " on" : "")} onClick={() => onTab("period")}>Period</button>
    <button role="tab" aria-selected={tab === "history"} className={"segbtn" + (tab === "history" ? " on" : "")} onClick={() => onTab("history")}>History</button>
  </div>);
}

/* ================================ History ================================
   Every saved month (or week, or day) side by side: metrics down the side, periods across,
   oldest on the left. A month or week with nothing uploaded shows as a gap. */
const HIST_WINDOWS = {
  month: [[3, "Last 3 months"], [6, "Last 6 months"], [12, "Last 12 months"], [0, "All months"]],
  week: [[13, "Last 13 weeks"], [26, "Last 26 weeks"], [52, "Last 52 weeks"], [0, "All weeks"]],
  day: [[14, "Last 14 days"], [31, "Last 31 days"], [90, "Last 90 days"], [0, "All days"]],
  custom: [[0, "All"]],
};
function historyColumns(periods, type, win) {
  const list = (periods || []).filter((p) => periodType(p) === type && p.range?.from && p.range?.to);
  if (!list.length) return [];
  if (type === "custom") return list.slice(0, win || list.length).reverse().map((p) => ({ key: p.id, from: p.range.from, to: p.range.to, label: shortLabel("custom", p.range.from, p.range.to), full: p.label, period: p }));
  const newest = list[0].range.from, oldest = list[list.length - 1].range.from;
  const count = win || Math.min(400, stepsBetween(type, oldest, newest) + 1);
  const cols = [];
  for (let i = 0; i < count; i++) {
    const r = stepRange(type, newest, -i);
    cols.push({ key: r.from, ...r, label: shortLabel(type, r.from, r.to), full: periodLabel(type, r.from, r.to), period: list.find((x) => x.range.from === r.from && x.range.to === r.to) || null });
  }
  return cols.reverse();
}
function slaCells(cols, key, name) {
  const cells = cols.map((c) => {
    const r = (c.period?.scopes?.[key]?.tasks || []).find((t) => t.name === name); if (!r) return null;
    const rec = n(r.received), ach = n(r.achieved); if (!(isNum(rec) && rec > 0 && isNum(ach))) return null;
    const pc = ach / rec; return { shown: Math.round(pc * 100) + "%", value: pc, st: pc >= 0.9 ? "ok" : pc >= 0.7 ? "warn" : "miss", reason: r.comment || "" };
  });
  return cells.some(Boolean) ? { key: key + "::" + name, name, cells, rule: { dir: "higher", good: 0.9, fair: 0.7, fmt: FMT.PERCENT }, fmt: FMT.PERCENT, isTxt: false, target: "≥ 90%" } : null;
}
function historyMatrix(client, cols) {
  const overall = cols.map((c) => {
    if (!c.period) return null; const sc = clientScorecard(client, c.period, null);
    return sc.pct == null ? null : { shown: sc.pct + "%", value: sc.pct / 100, st: sc.band || "info" };
  });
  const groups = [];
  for (const sc of (client?.scopes || []).map((k) => SCOPE_BY_KEY[k]).filter(Boolean)) {
    const rows = [];
    if (sc.criteria) for (const crit of sc.criteria) {
      const rule = effectiveRule(client, sc.key, crit), isTxt = ["status", "yesno", "text"].includes(crit.type);
      const cells = cols.map((c) => {
        const cell = c.period?.scopes?.[sc.key]?.crit?.[crit.key]; if (!cell || cell.value == null || cell.value === "") return null;
        const st = critStatus(crit, cell, rule);
        return { shown: critDisplay(crit, cell), value: isNum(cell.value) ? cell.value : null, st: st === "great" ? "ok" : st, reason: cell.reason || "", typed: cell.method === "manual" };
      });
      if (cells.some(Boolean)) rows.push({ key: sc.key + "." + crit.key, name: crit.name, cells, rule, fmt: critFmt(crit), isTxt, target: isTxt ? (crit.std || "") : targetText(rule) });
    }
    if (sc.kind === "tasks") for (const nm of sc.tasks || []) { const r = slaCells(cols, sc.key, nm); if (r) rows.push(r); }
    if (rows.length) groups.push({ key: sc.key, name: sc.name, rows });
  }
  const slaNames = []; for (const c of cols) for (const r of c.period?.scopes?.sla?.tasks || []) if (r.name && !slaNames.includes(r.name)) slaNames.push(r.name);
  const slaRows = slaNames.map((nm) => slaCells(cols, "sla", nm)).filter(Boolean);
  if (slaRows.length) groups.push({ key: "sla", name: "Service Level Agreement", rows: slaRows });
  return { overall, groups };
}
// newest value against the one before it, in points for percentages
function historyChange(row, cols) {
  const idx = []; row.cells.forEach((c, i) => { if (c && isNum(c.value)) idx.push(i); });
  if (idx.length < 2 || row.isTxt) return null;
  const a = row.cells[idx[idx.length - 2]], b = row.cells[idx[idx.length - 1]], d = b.value - a.value;
  const fmt = row.fmt, abs = Math.abs(d);
  const txt = fmt === FMT.PERCENT ? (abs * 100).toFixed(abs * 100 >= 10 ? 0 : 1) + " pts" : fmt === FMT.DAYS ? Math.round(abs) + (Math.round(abs) === 1 ? " day" : " days") : Math.round(abs).toLocaleString();
  return { d, text: d === 0 ? "No change" : (d > 0 ? "▲ " : "▼ ") + txt, cls: d === 0 ? "flat" : deltaCls(d, row.rule || { dir: "info" }), vs: cols[idx[idx.length - 2]].full };
}
const safeFileName = (s) => String(s || "client").replace(/[\\/:*?"<>|]+/g, " ").trim().slice(0, 80);
function downloadHistory(client, cols, matrix, type) {
  const wide = [["Scope", "Metric", "Target", ...cols.map((c) => c.full + (c.period ? "" : " (not uploaded)"))]];
  wide.push(["", "Overall score", "≥ 90%", ...matrix.overall.map((c) => (c ? c.shown : ""))]);
  for (const g of matrix.groups) for (const r of g.rows) wide.push([g.name, r.name, r.target, ...r.cells.map((c) => (c ? c.shown : ""))]);
  const long = [["Period type", "From", "To", "Period", "Scope", "Metric", "Value", "Unit", "Shown as", "Target", "Status", "Source", "Reason"]];
  for (const c of cols) if (c.period) for (const h of historyRows(client, c.period)) long.push([TYPE_NAME[type], c.from, c.to, c.full, h.scope, h.metric, h.value == null ? "" : h.value, h.unit, h.shown, h.target, h.status, h.source, h.reason]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(wide), "History");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(long), "All rows");
  XLSX.writeFile(wb, safeFileName(client.name) + " - " + TYPE_NAME[type] + " history.xlsx");
}
function ClientHistory({ client, onOpenPeriod }) {
  const periods = client?.periods || [];
  const typesHere = ["month", "week", "day", "custom"].filter((t) => periods.some((p) => periodType(p) === t));
  const [type, setType] = useState(typesHere[0] || "month");
  const [win, setWin] = useState(HIST_WINDOWS[typesHere[0] || "month"][0][0]);
  const [open, setOpen] = useState(null);
  useEffect(() => { if (typesHere.length && !typesHere.includes(type)) { setType(typesHere[0]); setWin(HIST_WINDOWS[typesHere[0]][0][0]); } }, [typesHere.join(",")]);
  const cols = useMemo(() => historyColumns(periods, type, win), [periods, type, win]);
  const matrix = useMemo(() => historyMatrix(client, cols), [client, cols]);
  if (!periods.length) return <div className="card" style={{ color: "var(--ink-3)" }}>Nothing saved yet. Pick a month, week or day under <b>Period</b> and upload its reports. Every period you upload is kept and shows up here.</div>;
  const setT = (t) => { setType(t); setWin(HIST_WINDOWS[t][0][0]); setOpen(null); };
  const uploaded = cols.filter((c) => c.period).length, missing = cols.length - uploaded;
  const span = 2 + cols.length;
  const head = (c) => (c.period && onOpenPeriod ? <button className="hist-ph" title={"Open " + c.full} onClick={() => onOpenPeriod(c.period.id)}>{c.label}</button> : <span title={c.full}>{c.label}</span>);
  const cellTitle = (c, cell) => cell ? `${c.full}: ${cell.shown}${cell.typed ? " (typed in)" : ""}${cell.reason ? ". Reason: " + cell.reason : ""}` : c.period ? c.full + ": no value" : c.full + ": not uploaded";
  const overallRow = { name: "Overall score", cells: matrix.overall, rule: { dir: "higher", good: 0.9, fair: 0.8, fmt: FMT.PERCENT }, fmt: FMT.PERCENT, isTxt: false, target: "≥ 90%", key: "overall" };
  const rowView = (r, extra) => {
    const ch = historyChange(r, cols);
    return (
      <React.Fragment key={r.key}>
        <tr className={"hist-r" + (open === r.key ? " on" : "") + (extra || "")} onClick={() => setOpen(open === r.key ? null : r.key)} title="Click for a trend chart">
          <td className="hm"><span className="hist-name">{r.name}</span>{r.target && <span className="hist-target">Target {r.target}</span>}</td>
          {r.cells.map((cell, i) => <td key={i} className={"hist-c " + (cols[i].period ? (cell ? cell.st : "none") : "gap")} title={cellTitle(cols[i], cell)}>{cell ? cell.shown : cols[i].period ? "—" : ""}</td>)}
          <td className={"hist-chg " + (ch ? ch.cls : "flat")} title={ch ? "Compared with " + ch.vs : ""}>{ch ? ch.text : ""}</td>
        </tr>
        {open === r.key && <tr className="hist-chart-row"><td colSpan={span}><TrendChart row={r} cols={cols} /></td></tr>}
      </React.Fragment>
    );
  };
  return (
    <div>
      <div className="hist-ctl">
        {typesHere.length > 1 && <div className="seg">{typesHere.map((t) => <button key={t} className={"segbtn" + (type === t ? " on" : "")} onClick={() => setT(t)}>{TYPE_NAME[t]}</button>)}</div>}
        {HIST_WINDOWS[type].length > 1 && <select value={win} onChange={(e) => { setWin(Number(e.target.value)); setOpen(null); }} aria-label="How far back">{HIST_WINDOWS[type].map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select>}
        <span className="hist-note">{uploaded} {TYPE_UNIT[type]}{uploaded === 1 ? "" : "s"} saved{missing > 0 ? `, ${missing} not uploaded` : ""}. Click a metric for its trend.</span>
        <span style={{ flex: 1 }} />
        <button className="btn sm" onClick={() => downloadHistory(client, cols, matrix, type)}>Download Excel</button>
      </div>
      <div className="hist-wrap">
        <table className="hist">
          <thead><tr>
            <th className="hm">Metric</th>
            {cols.map((c) => <th key={c.key} className={c.period ? "" : "gap"}>{head(c)}{!c.period && <span className="hist-miss">Not uploaded</span>}</th>)}
            <th className="hist-chg-h">Change</th>
          </tr></thead>
          <tbody>
            {rowView(overallRow, " hist-overall")}
            {matrix.groups.map((g) => (
              <React.Fragment key={g.key}>
                <tr className="hist-g"><td className="hm" colSpan={1}><span style={{ display: "inline-flex", alignItems: "center", gap: 7 }}><ScopeTag k={g.key} />{g.name}</span></td><td colSpan={span - 1} /></tr>
                {g.rows.map((r) => rowView(r))}
              </React.Fragment>
            ))}
          </tbody>
        </table>
      </div>
      {matrix.groups.length === 0 && <div className="card" style={{ color: "var(--ink-3)", marginTop: 12 }}>No numbers in these {TYPE_UNIT[type]}s yet.</div>}
    </div>
  );
}
function niceTicks(lo, hi, count = 4) {
  if (!(hi > lo)) { const d = Math.abs(lo) || 1; lo -= d * 0.1; hi += d * 0.1; }
  const raw = (hi - lo) / count, mag = Math.pow(10, Math.floor(Math.log10(raw))), norm = raw / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  const out = []; for (let v = Math.floor(lo / step) * step; v <= Math.ceil(hi / step) * step + step / 2; v += step) out.push(+v.toFixed(10));
  return out;
}
// One metric over time: a 2px line with dots, the target as a thin reference line, and a hover readout.
function TrendChart({ row, cols }) {
  const wrapRef = useRef(null);
  const [w, setW] = useState(640);
  const [hover, setHover] = useState(null);
  // as wide as the visible part of the table, and pinned to its left edge when the table scrolls sideways
  useEffect(() => {
    const el = wrapRef.current; if (!el) return;
    const host = el.closest(".hist-wrap") || el.parentElement || el;
    const fit = () => setW(Math.max(300, Math.round(host.clientWidth - 30)));
    fit(); let ro = null;
    try { ro = new ResizeObserver(fit); ro.observe(host); } catch { window.addEventListener("resize", fit); }
    return () => { if (ro) ro.disconnect(); else window.removeEventListener("resize", fit); };
  }, []);
  const vals = row.cells.filter((c) => c && isNum(c.value)).map((c) => c.value);
  if (!vals.length) return <div ref={wrapRef} className="trend-none">This metric is a word (such as Met or Not Met), so there is no line to draw. The values are in the row above.</div>;
  const rule = row.rule || {};
  const targets = rule.dir === "range" ? [rule.min, rule.max].filter(isNum) : rule.dir === "higher" || rule.dir === "lower" ? [rule.good].filter(isNum) : [];
  const ticks = niceTicks(Math.min(...vals, ...targets), Math.max(...vals, ...targets), 4);
  const lo = ticks[0], hi = ticks[ticks.length - 1];
  const H = 190, padL = 68, padR = 24, padT = 16, padB = 30, iw = w - padL - padR, ih = H - padT - padB;
  const x = (i) => padL + (cols.length === 1 ? iw / 2 : (i * iw) / (cols.length - 1));
  const y = (v) => padT + (1 - (v - lo) / (hi - lo || 1)) * ih;
  let d = "", pen = false;
  row.cells.forEach((c, i) => { if (c && isNum(c.value)) { d += (pen ? "L" : "M") + x(i).toFixed(1) + " " + y(c.value).toFixed(1) + " "; pen = true; } else pen = false; });
  const every = Math.max(1, Math.ceil(cols.length / Math.max(2, Math.floor(iw / 72))));
  const nearest = (px) => { let best = 0, bd = Infinity; cols.forEach((_, i) => { const dd = Math.abs(x(i) - px); if (dd < bd) { bd = dd; best = i; } }); return best; };
  const fv = (v) => fmtVal(v, row.fmt);
  const hc = hover == null ? null : row.cells[hover];
  return (
    <div className="trend" ref={wrapRef} style={{ width: w }}>
      <div className="trend-h"><b>{row.name}</b>{row.target && <span>Target {row.target}</span>}</div>
      <svg width={w} height={H} role="img" aria-label={`${row.name} over ${cols.length} periods`} tabIndex={0} style={{ display: "block", outline: "none" }}
        onMouseMove={(e) => { const r = e.currentTarget.getBoundingClientRect(); setHover(nearest(e.clientX - r.left)); }} onMouseLeave={() => setHover(null)}
        onFocus={() => setHover(cols.length - 1)} onBlur={() => setHover(null)}
        onKeyDown={(e) => { if (e.key === "ArrowLeft") setHover((h) => Math.max(0, (h ?? cols.length - 1) - 1)); if (e.key === "ArrowRight") setHover((h) => Math.min(cols.length - 1, (h ?? 0) + 1)); }}>
        {ticks.map((t) => <g key={t}><line x1={padL} x2={w - padR} y1={y(t)} y2={y(t)} style={{ stroke: "var(--line-2)" }} strokeWidth="1" /><text x={padL - 8} y={y(t) + 4} textAnchor="end" style={{ fill: "var(--ink-3)", fontSize: 11 }}>{fv(t)}</text></g>)}
        {targets.map((t) => <g key={"t" + t}><line x1={padL} x2={w - padR} y1={y(t)} y2={y(t)} style={{ stroke: "var(--ink-3)" }} strokeWidth="1" opacity=".75" /><text x={w - padR} y={y(t) - 5} textAnchor="end" style={{ fill: "var(--ink-2)", fontSize: 10.5, fontWeight: 600 }}>Target {fv(t)}</text></g>)}
        {cols.map((c, i) => (i % every === 0 || i === cols.length - 1) && <text key={c.key} x={x(i)} y={H - 9} textAnchor={cols.length > 1 && i === cols.length - 1 ? "end" : cols.length > 1 && i === 0 ? "start" : "middle"} style={{ fill: "var(--ink-3)", fontSize: 10.5 }}>{c.label}</text>)}
        {hover != null && <line x1={x(hover)} x2={x(hover)} y1={padT} y2={padT + ih} style={{ stroke: "var(--ink-3)" }} strokeWidth="1" />}
        <path d={d} fill="none" style={{ stroke: "var(--steel)" }} strokeWidth="2" strokeLinejoin="round" strokeLinecap="round" />
        {row.cells.map((c, i) => c && isNum(c.value) && <circle key={i} cx={x(i)} cy={y(c.value)} r={hover === i ? 5.5 : 4.5} style={{ fill: "var(--steel)", stroke: "#fff" }} strokeWidth="2" />)}
      </svg>
      {hover != null && <div className="trend-tip" style={{ left: Math.min(Math.max(x(hover), 70), w - 70), top: (hc && isNum(hc.value) ? y(hc.value) : padT + ih / 2) + 18 }}>
        <b>{hc ? hc.shown : cols[hover].period ? "No value" : "Not uploaded"}</b>
        <span>{cols[hover].full}{hc && hc.st && hc.st !== "info" ? " · " + stLabel(hc.st) : ""}</span>
        {hc?.reason && <span className="trend-why">{hc.reason}</span>}
      </div>}
    </div>
  );
}

function JoinKeyPicker({ req, onClose, onApprove }) {
  const [sel, setSel] = useState(() => new Set(req.fields || []));
  const toggle = (id) => setSel((s) => { const n2 = new Set(s); n2.has(id) ? n2.delete(id) : n2.add(id); return n2; });
  const labelFor = { patient: "Patient", dos: "Date of Service", cpt: "CPT", payer: "Payer", provider: "Provider", billed: "Billed amount" };
  return (<Modal onClose={onClose} title="Confirm how to match these files" desc="These two files don't share a claim identifier (Claim ID, Account No, Encounter No…). To match them, I'll combine the fields below into one key. Confirm the combination, or adjust it.">
    <div className="field"><label>Fields to combine into a match key</label>
      <div style={{ display: "flex", flexDirection: "column", gap: 6, marginTop: 4 }}>
        {req.fields.map((id) => (
          <label key={id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, cursor: "pointer" }}>
            <input type="checkbox" checked={sel.has(id)} onChange={() => toggle(id)} /> {labelFor[id] || id}
          </label>
        ))}
      </div>
    </div>
    <div className="hint" style={{ fontSize: 11.5, color: "var(--ink-3)" }}>I'll use this combination to match rows across the two files. Once you approve it, I'll reuse it for this client and only ask again if the columns change.</div>
    <Actions onClose={onClose} onOk={() => onApprove([...sel])} okLabel="Approve match key" okDisabled={sel.size === 0} />
  </Modal>);
}
/* ===================== OneDrive: upload the reports, save the summary ===================== */
// Each person signs in to Microsoft with their own Purview work account (a pop-up), and OneDrive is read
// and written through Microsoft Graph as that person, so everything is done as, and recorded against,
// whoever clicked. Drishti only ever reads the team's report files. The one file it writes is its own
// "Drishti summary - <client>.xlsx", changed in place through the Excel workbook API, never replaced by an
// uploaded copy. The Google Sheet backend is untouched: the client's folder and the summary bookkeeping ride
// along in the client's own settings, which the backend already saves as they are.
// Every upload comes from OneDrive once it is switched on: the person opens My files of the team's shared
// OneDrive account (MS_ACCOUNT below), or the folder a link in window.DRISHTI_ONEDRIVE_FOLDER opens, and
// picks the files. Nothing is uploaded from the computer then.
const MS_HOME = "https://purview-healthcare.github.io/Drishti/";
// Drishti has its own Microsoft app registration ("Purview Drishti"). Its Application (client) ID is set in
// config.js as window.DRISHTI_MS_CLIENT_ID; until it is, OneDrive stays switched off. The tenant is
// purview-healthcare.com unless config.js sets window.DRISHTI_MS_TENANT_ID. Both are public identifiers.
const MS_TENANT = "94be55e5-2194-40e4-9d98-1dcc7e91c268";
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const cfgText = (k) => (typeof window !== "undefined" && window[k] != null ? String(window[k]).trim() : "");
// Everyone signs in to Microsoft as the team's shared account and uploads from its OneDrive. config.js can
// name another one with window.DRISHTI_MS_ACCOUNT, or set it to "" so each person uses their own.
const MS_ACCOUNT = "Purviewautomate@purview-healthcare.com";
const odAccountShown = () => (typeof window !== "undefined" && window.DRISHTI_MS_ACCOUNT != null ? cfgText("DRISHTI_MS_ACCOUNT") : MS_ACCOUNT);
const MS_CFG = {
  get clientId() { return cfgText("DRISHTI_MS_CLIENT_ID"); },
  get tenantId() { const t = cfgText("DRISHTI_MS_TENANT_ID"); return GUID_RE.test(t) ? t : MS_TENANT; },
  get rootLink() { return cfgText("DRISHTI_ONEDRIVE_FOLDER"); },
  get account() { return odAccountShown().toLowerCase(); },
  scopes: ["User.Read", "Files.ReadWrite.All"],
};
const msConfigured = () => GUID_RE.test(MS_CFG.clientId);
const GRAPH = "https://graph.microsoft.com/v1.0";
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const SUMMARY_PREFIX = "Drishti summary - ";
const OD_TABS = { month: "Monthly", week: "Weekly" };
const OD_HOWTO = "How to read this";
const ODV = 2; // layout version of the summary file; a higher number rewrites every tab once
const enc = encodeURIComponent;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const msMock = () => (typeof window !== "undefined" && window.__DRISHTI_MS_MOCK) || null;
// Microsoft sends a sign-in back only to an address registered for the app, so this works on the live
// Drishti site (and in tests, with the pretend Microsoft), not on a copy hosted anywhere else.
function msRedirectUri() {
  try { const here = window.location.origin + window.location.pathname.replace(/index\.html?$/i, ""); return here.toLowerCase() === MS_HOME.toLowerCase() ? MS_HOME : here; }
  catch { return MS_HOME; }
}
function msSupported() {
  if (!msConfigured()) return false;
  if (msMock()) return true;
  try { return !!(window.msal && window.msal.PublicClientApplication) && window.location.origin === new URL(MS_HOME).origin; } catch { return false; }
}
let MSAL_P = null;
function msalApp() {
  if (!MSAL_P) {
    const app = new window.msal.PublicClientApplication({
      auth: { clientId: MS_CFG.clientId, authority: "https://login.microsoftonline.com/" + MS_CFG.tenantId, redirectUri: msRedirectUri() },
      cache: { cacheLocation: "localStorage" },
      system: { popupBridgeTimeout: 300000 }, // five minutes, time enough for an Authenticator code
    });
    const p = app.initialize().then(() => app);
    p.catch(() => { if (MSAL_P === p) MSAL_P = null; });
    MSAL_P = p;
  }
  return MSAL_P;
}
function msErr(kind, message, status) { const e = new Error(message); e.kind = kind; e.status = status || 0; return e; }
function graphError(status, body) {
  const code = (body && body.error && body.error.code) || "";
  if (status === 401) return msErr("auth", "Microsoft asked you to sign in again. Click the button once more and sign in.", status);
  if (status === 403) return msErr("denied", MS_CFG.account ? `OneDrive did not let this account open that. Check that Microsoft is signed in as ${odAccountShown()}, then try again.` : "You do not have access to this in OneDrive yet. Ask whoever looks after the Drishti folder in OneDrive to share it with you, with Can edit.", status);
  if (status === 404) return msErr("missing", "OneDrive could not find the folder or file. It may have been deleted, or moved somewhere you cannot see. Open the folders again and pick it.", status);
  if (status === 409 || status === 423) return msErr("busy", "The summary file is busy, or open in the Excel desktop app. Close it there, wait a minute and try again.", status);
  if (status === 429 || status === 503 || status === 504) return msErr("throttled", "Microsoft is busy right now. Wait a minute and try again.", status);
  if (status >= 500) return msErr("server", "OneDrive had a problem on its side. Wait a minute and try again.", status);
  return msErr("other", "OneDrive turned this down (" + status + (code ? ", " + code : "") + "). Try again, and if it keeps happening send it through Feedback.", status);
}
// one plain sentence for anything that can go wrong with Microsoft sign in or OneDrive
function msalMessage(e) {
  if (e && e.kind) return e.message;
  const code = String((e && (e.errorCode || e.code)) || "");
  const text = String((e && (e.errorMessage || e.message)) || "");
  if (code === "user_cancelled") return "Microsoft sign in was closed before it finished. Click again to sign in.";
  if (code === "popup_window_error" || code === "empty_window_error") return "Your browser blocked the Microsoft sign in window. Allow pop-ups for this site, then click again.";
  if (code === "timed_out") return "Microsoft sign in took too long. Click again to retry.";
  if (code === "interaction_in_progress") return "A Microsoft sign in window is already open. Finish it there, or close it and click again.";
  if (/AADSTS65001|AADSTS90094|AADSTS90008|consent|approval/i.test(text)) return "Microsoft needs an admin to approve Drishti for Purview first. Ask your Microsoft 365 admin to open Entra ID, App registrations, Purview Drishti, API permissions, and click Grant admin consent.";
  if (/AADSTS50020|AADSTS50034|AADSTS50177|not exist in tenant/i.test(text)) return MS_CFG.account ? `That Microsoft account is not part of Purview's Microsoft 365. Click again and sign in as ${odAccountShown()}.` : "That Microsoft account is not part of the Purview directory that holds the files. Sign in with your Purview work account, or ask the admin to invite your account as a guest.";
  if (/AADSTS700016|AADSTS90002|application .* not found/i.test(text)) return "Microsoft does not know the app ID in Drishti's settings. Check that config.js has the Application (client) ID of the Purview Drishti app registration.";
  if (/AADSTS50011|AADSTS50194|reply url|redirect/i.test(text)) return "Microsoft did not accept this page's address. Drishti's address has to be added to the Purview Drishti app registration as a single-page application redirect URI.";
  if (/AADSTS50076|AADSTS50079|multi-?factor|\bmfa\b/i.test(text)) return "Microsoft needs you to finish multi-factor sign in. Click again and complete it in the window.";
  return "Microsoft sign in did not finish" + (code ? " (" + code + ")" : "") + ". Try again.";
}
const RETRY_STATUS = new Set([429, 503, 504]);
function retryDelay(r, attempt) {
  const m = msMock(); if (m && m.retryMs != null) return m.retryMs;
  const ra = Number(r && r.headers && r.headers.get ? r.headers.get("Retry-After") : 0);
  return Math.min(20000, ra > 0 ? ra * 1000 : 1200 * Math.pow(2, attempt));
}
const msx = {
  hint: "", // the Drishti user's email, so the matching Microsoft account is used on a shared computer
  // The account Drishti uses: the one picked on this computer with Switch, otherwise the one whose email
  // is the Drishti user's own. Never somebody else's, even if another app on this site signed them in.
  async account(hint) {
    const m = msMock(); if (m) return m.account ? m.account(hint) : null;
    const app = await msalApp(); const active = app.getActiveAccount(); if (active) return active;
    const h = String(hint || "").toLowerCase();
    return (h && app.getAllAccounts().find((a) => String(a.username || "").toLowerCase() === h)) || null;
  },
  async signIn(hint, pick) {
    const m = msMock(); if (m) return m.signIn(hint, pick);
    const app = await msalApp();
    const r = await app.loginPopup({ scopes: MS_CFG.scopes, loginHint: pick ? undefined : hint || undefined, prompt: pick ? "select_account" : undefined, overrideInteractionInProgress: true });
    app.setActiveAccount(r.account); return r.account;
  },
  async token() {
    const m = msMock(); if (m) return m.token();
    const app = await msalApp(); const acct = await msx.account(msx.hint);
    if (!acct) throw msErr("signin", "Sign in to Microsoft first. Click the button again.");
    try { return (await app.acquireTokenSilent({ scopes: MS_CFG.scopes, account: acct })).accessToken; }
    catch {
      // an expired session or a newly added permission needs the person once more
      const r = await app.acquireTokenPopup({ scopes: MS_CFG.scopes, account: acct, overrideInteractionInProgress: true });
      app.setActiveAccount(r.account); return r.accessToken;
    }
  },
  // Signing out of Drishti forgets which Microsoft account Drishti was using on this computer. It does not
  // log out or clear MSAL's cache: the account record is shared by every app on this site, and clearing it
  // would sign the person out of Brahmastra too. The next person is never given this account (see account).
  async forget() {
    const m = msMock(); if (m) { if (m.forget) await m.forget(); return; }
    const app = await msalApp(); app.setActiveAccount(null);
  },
  async request(url, opts = {}, auth = true) {
    const m = msMock(); const doFetch = (m && m.fetch) || ((u, o) => fetch(u, o));
    for (let attempt = 0; ; attempt++) {
      const headers = { ...(opts.headers || {}) };
      if (auth) headers.Authorization = "Bearer " + (await msx.token());
      let r;
      try { r = await doFetch(url, { method: opts.method || "GET", headers, body: opts.body }); }
      catch { throw msErr("network", "Could not reach Microsoft. Check your connection and try again."); }
      if (r.ok) return r;
      if (RETRY_STATUS.has(r.status) && attempt < 4) { await sleep(retryDelay(r, attempt)); continue; }
      let j = null; try { j = await r.json(); } catch { j = null; }
      throw graphError(r.status, j);
    }
  },
  async graph(path, opts = {}) {
    const r = await msx.request(GRAPH + path, opts, true);
    if (r.status === 204 || r.status === 202) return null;
    const t = await r.text(); return t ? JSON.parse(t) : null;
  },
  // OneDrive's short-lived download address is already signed, so it is fetched with no Authorization
  // header, which is Microsoft's advice for apps running in a browser
  async download(url) { const r = await msx.request(url, {}, false); return new Uint8Array(await r.arrayBuffer()); },
};
// Many small Excel changes in one trip: up to 20 per request, run in order, retried from where Microsoft
// asked to slow down.
async function graphBatch(reqs, sid) {
  for (let i = 0; i < reqs.length; i += 20) {
    let part = reqs.slice(i, i + 20);
    for (let attempt = 0; ; attempt++) {
      const requests = part.map((q, k) => ({ id: String(k + 1), method: q.method, url: q.url, headers: { ...(q.body !== undefined ? { "Content-Type": "application/json" } : {}), ...(sid ? { "workbook-session-id": sid } : {}) }, ...(q.body !== undefined ? { body: q.body } : {}), ...(k ? { dependsOn: [String(k)] } : {}) }));
      const res = await msx.graph("/$batch", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ requests }) });
      const byId = new Map(((res && res.responses) || []).map((x) => [String(x.id), x]));
      let failAt = -1, fail = null;
      for (let k = 0; k < part.length; k++) { const x = byId.get(String(k + 1)); if (x && x.status >= 200 && x.status < 300) continue; failAt = k; fail = x || null; break; }
      if (failAt < 0) break;
      if (fail && RETRY_STATUS.has(fail.status) && attempt < 4) { part = part.slice(failAt); await sleep(retryDelay(null, attempt)); continue; }
      throw graphError(fail ? fail.status : 0, fail && fail.body);
    }
  }
}
function shareIdOf(link) {
  const b64 = btoa(unescape(encodeURIComponent(String(link).trim())));
  return "u!" + b64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
}
const odItem = (od, id) => `/drives/${enc(od.driveId)}/items/${enc(id)}`;
const personName = (me) => String((me && (me.name || me.firstName)) || displayName(me) || "").trim();
function odPathOf(it) {
  const pr = (it && it.parentReference) || {};
  let p = String(pr.path || "").split("root:")[1] || "";
  try { p = decodeURIComponent(p); } catch {}
  return p.split("/").filter(Boolean).concat(it.name).join(" › ");
}
// a share link pasted by a person: it must open a folder, the client's own
async function odResolveFolder(link) {
  const s = String(link || "").trim();
  if (!/^https:\/\//i.test(s)) throw msErr("input", "Paste the whole link, starting with https://");
  const it = await msx.graph("/shares/" + shareIdOf(s) + "/driveItem");
  if (!it || !it.folder) throw msErr("input", `That link opens ${it && it.name ? `"${it.name}", which is a file` : "a file"}. Copy the link of the client's folder instead.`);
  const pr = it.parentReference || {};
  if (!pr.driveId) throw msErr("input", "OneDrive did not say which drive this folder is in. Copy the link again from the folder itself and paste it here.");
  return { driveId: pr.driveId, itemId: it.id, name: it.name, webUrl: it.webUrl || s, path: odPathOf(it) };
}
// Where browsing starts: the folder a link in config.js opens, or else My files of the signed-in account.
let OD_ROOT = null;
function odRoot() {
  const link = MS_CFG.rootLink, key = link || "me";
  if (!OD_ROOT || OD_ROOT.key !== key) {
    const p = (link ? odResolveFolder(link) : odMyFiles()).catch((e) => {
      if (OD_ROOT && OD_ROOT.p === p) OD_ROOT = null;
      if (e.status === 403) throw msErr("denied", "You do not have access to that OneDrive folder yet. Ask whoever looks after it to share it with you, with Can edit.", 403);
      if (e.status === 404 || e.kind === "input") throw msErr("setup", "Drishti's OneDrive setting does not open a folder. Ask the Admin to check it.", e.status);
      throw e;
    });
    OD_ROOT = { key, p };
  }
  return OD_ROOT.p;
}
async function odMyFiles() {
  const r = await msx.graph("/me/drive/root?$select=id,name,webUrl,parentReference");
  let driveId = r && r.parentReference && r.parentReference.driveId;
  if (!driveId) { const d = await msx.graph("/me/drive?$select=id"); driveId = d && d.id; }
  if (!r || !r.id || !driveId) throw msErr("setup", "OneDrive did not open for this account. Try again in a minute.");
  return { driveId, itemId: r.id, name: "My files", webUrl: r.webUrl || "", path: "My files" };
}
// The Microsoft account Drishti signs in with: the shared one in config.js, or the person's own Purview email.
const OD_ACCT_EVENT = "drishti-ms-account";
const odHint = (me) => MS_CFG.account || (me && me.username) || "";
const odRightAccount = (a) => !MS_CFG.account || String((a && a.username) || "").toLowerCase() === MS_CFG.account;
const odAnnounce = (a) => { try { window.dispatchEvent(new CustomEvent(OD_ACCT_EVENT, { detail: a || null })); } catch { /* only tells the bar */ } };
async function odCheckAccount(a) {
  if (!odRightAccount(a)) {
    await msx.forget(); odAnnounce(null);
    const who = String((a && a.username) || "").trim(), want = odAccountShown();
    throw msErr("account", who ? `Microsoft signed in ${who}, but Drishti uploads only from the OneDrive of ${want}. Click again and sign in as ${want}.` : `Drishti uploads only from the OneDrive of ${want}. Click again and sign in as ${want}.`);
  }
  return a;
}
async function odEnsure(hint, onStep) {
  if (hint) msx.hint = hint;
  let a = await msx.account(hint);
  // another Microsoft account left active on this computer is set aside, not used
  if (a && !odRightAccount(a)) { await msx.forget(); a = await msx.account(hint); if (a && !odRightAccount(a)) a = null; }
  if (!a) { if (onStep) onStep("Signing in to Microsoft…"); a = await msx.signIn(hint, false); }
  await odCheckAccount(a);
  await msx.token();
  odAnnounce(a);
  return a;
}
async function odChildren(od, id) {
  let url = odItem(od, id) + "/children?$top=200"; const out = [];
  for (let guard = 0; url && guard < 50; guard++) {
    const j = await msx.graph(url);
    out.push(...((j && j.value) || []));
    const nx = j && j["@odata.nextLink"]; url = nx ? String(nx).replace(GRAPH, "") : null;
  }
  return out;
}
const isReportItem = (x) => !!x && !!x.file && BULK_IN.test(x.name || "") && !isJunkName(x.name || "") && !String(x.name).startsWith(SUMMARY_PREFIX);

/* Month and week folders. Months may be written "October 2026", "Oct 2026", "oct-2026", "10-2026",
   "2026-10" or "Oct'26", weeks "Week 1", "Wk 1" or "W1". Drishti uses this only to tell the client's own
   folder apart from the month and week folders inside it, so the summary file goes in the client's folder.
   A week runs Monday to Friday and belongs to the month its Monday is in; Week 1 is the week of the
   month's first Monday. */
const MONTH_WORDS = [["january", "jan"], ["february", "feb"], ["march", "mar"], ["april", "apr"], ["may"], ["june", "jun"], ["july", "jul"], ["august", "aug"], ["september", "sep", "sept"], ["october", "oct"], ["november", "nov"], ["december", "dec"]];
const monthOfWord = (w) => MONTH_WORDS.findIndex((a) => a.includes(w));
function nameTokens(name) {
  return String(name || "").toLowerCase().replace(/([a-z])(\d)/g, "$1 $2").replace(/(\d)([a-z])/g, "$1 $2").replace(/[^a-z0-9]+/g, " ").trim().split(" ").filter(Boolean);
}
function weekInfo(fromYmd) {
  const mon = parseYmd(fromYmd), fri = addDays(mon, 4);
  return { y: mon.getFullYear(), m: mon.getMonth(), n: Math.floor((mon.getDate() - 1) / 7) + 1, mon, fri };
}
const weekDatesText = (wk) => (wk.mon.getMonth() === wk.fri.getMonth() ? `${wk.mon.getDate()} to ${wk.fri.getDate()} ${MSHORT[wk.fri.getMonth()]}` : `${wk.mon.getDate()} ${MSHORT[wk.mon.getMonth()]} to ${wk.fri.getDate()} ${MSHORT[wk.fri.getMonth()]}`);
function weekNumberOf(name) {
  const t = nameTokens(name);
  for (let i = 0; i < t.length - 1; i++) {
    if (!["week", "wk", "w", "weeks", "wks"].includes(t[i])) continue;
    if (/^\d{1,2}$/.test(t[i + 1])) return Number(t[i + 1]);
    if (["no", "number", "num"].includes(t[i + 1]) && /^\d{1,2}$/.test(t[i + 2] || "")) return Number(t[i + 2]);
  }
  return t.length === 1 && /^[1-6]$/.test(t[0]) ? Number(t[0]) : null;
}
// A folder named after a month or a week ("October 2026", "Oct'26", "10-2026", "Week 1", "Wk 2", "W3").
function isPeriodName(name) {
  const t = nameTokens(name);
  if (weekNumberOf(name) != null || t.some((x) => x === "week" || x === "weeks")) return true;
  const years = t.filter((x) => /^(\d{4}|\d{2})$/.test(x));
  if (t.some((x) => monthOfWord(x) >= 0) && years.length) return true;
  const nums = t.filter((x) => /^\d{1,4}$/.test(x)).map(Number);
  return nums.length === 2 && nums.some((n) => n >= 2000 && n <= 2100) && nums.some((n) => n >= 1 && n <= 12);
}
// The client's folder, from where report files were picked: above any month or week folders.
function odClientTrail(trail) {
  let k = trail.length;
  while (k > 1 && isPeriodName(trail[k - 1].name)) k--;
  return trail.slice(0, k);
}
async function odReportFiles(od, folderId) { return (await odChildren(od, folderId)).filter(isReportItem); }
async function odDownloadFiles(od, items, onProgress) {
  const out = new Array(items.length); let next = 0, done = 0;
  const worker = async () => {
    while (next < items.length) {
      const i = next++, it = items[i];
      let url = it["@microsoft.graph.downloadUrl"];
      if (!url) { const full = await msx.graph(odItem(od, it.id)); url = full && full["@microsoft.graph.downloadUrl"]; }
      if (!url) throw msErr("other", `OneDrive did not offer ${it.name} for download. Try again in a minute.`);
      const f = new File([await msx.download(url)], it.name, { lastModified: Date.parse(it.lastModifiedDateTime) || Date.now() });
      f._folder = ""; out[i] = f; done++; if (onProgress) onProgress(done, items.length);
    }
  };
  await Promise.all(Array.from({ length: Math.min(4, items.length) }, worker));
  return out;
}

/* ---- the summary file ---- */
const summaryFileName = (clientName) => SUMMARY_PREFIX + safeFileName(clientName) + ".xlsx";
const xlText = (s) => { const t = String(s == null ? "" : s); return /^[=+\-@]/.test(t) ? " " + t : t; };
const colLetter = (i) => { let s = "", x = i + 1; while (x > 0) { const r = (x - 1) % 26; s = String.fromCharCode(65 + r) + s; x = Math.floor((x - 1) / 26); } return s; };
const cellAddr = (c1, r1, c2, r2) => `${colLetter(c1)}${r1}:${colLetter(c2)}${r2}`;
const OD_STYLE = {
  ok: { fill: "#DFF3E7", font: "#155C3D", bold: true }, warn: { fill: "#FBEFD5", font: "#7A5209", bold: true }, miss: { fill: "#FBE1E1", font: "#8A2B2B", bold: true },
  gap: { fill: "#F1F3F6" }, "chg-up": { font: "#1F9D6B", bold: true }, "chg-down": { font: "#C23B3B", bold: true }, "chg-flat": { font: "#8A97AB", bold: false },
};
const OD_INK = "#16233B", OD_INK2 = "#4A5A72";
const odStyleOf = (st) => (st === "great" || st === "ok" ? "ok" : st === "warn" ? "warn" : st === "miss" ? "miss" : "none");
// the same number format the screens use: whole percentages without a decimal, others with one
function odNumFmt(fmt, value, whole) {
  if (fmt === FMT.PERCENT) return whole || (value * 100) % 1 === 0 ? "0%" : "0.0%";
  if (fmt === FMT.DAYS) return '[=1]0" day";0" days"';
  if (fmt === FMT.COUNT) return "#,##0";
  return "General";
}
function odColHead(kind, col) {
  if (kind === "week") { const wk = weekInfo(col.from); return `${MSHORT[wk.m]} ${wk.y} Week ${wk.n}\n${col.period ? weekDatesText(wk) : "not uploaded"}`; }
  return col.period ? col.label : `${col.label}\nnot uploaded`;
}
// One tab of the summary, worked out from exactly what the History tab shows: one row per KPI and SLA
// line, one column per month (or week) from the oldest to the newest, and the change since the last one.
// a month or week someone opened but never filled in is left out, or shown as not uploaded between two others
const periodHasValues = (p) => Object.values((p && p.scopes) || {}).some((s) => s && (Object.values(s.crit || {}).some((c) => c && c.value != null && c.value !== "") || (s.tasks || []).some((t) => t && ((t.received != null && t.received !== "") || (t.achieved != null && t.achieved !== "")))));
function odTabModel(client, kind, stampText) {
  const cols = historyColumns((client.periods || []).filter(periodHasValues), kind);
  const matrix = historyMatrix(client, cols);
  const rows = [];
  if (cols.length) rows.push({ key: "overall", scope: "All", name: "Overall score", formula: "Metrics on target ÷ metrics scored", target: "≥ 90%", cells: matrix.overall, rule: { dir: "higher", good: 0.9, fair: 0.8, fmt: FMT.PERCENT }, fmt: FMT.PERCENT, isTxt: false, whole: true });
  for (const g of matrix.groups) for (const r of g.rows) {
    const task = String(r.key).includes("::");
    let formula = "Achieved ÷ received";
    if (!task) { const [sk, ck] = String(r.key).split("."); const crit = critOf(sk, ck); formula = (crit && (crit.formula || crit.std)) || ""; }
    rows.push({ ...r, scope: g.name, formula, whole: task });
  }
  const head = ["Scope", "Metric", "Formula", "Target", ...cols.map((c) => odColHead(kind, c)), kind === "week" ? "Change vs last week" : "Change vs last month"];
  const colCells = cols.map((c, j) => rows.map((r) => {
    const cell = r.cells[j];
    if (!c.period) return { v: "", nf: "General", style: "gap" };
    if (!cell) return { v: "", nf: "General", style: "none" };
    if (r.isTxt || !isNum(cell.value)) return { v: xlText(cell.shown), nf: "@", style: odStyleOf(cell.st) };
    return { v: Math.round(cell.value * 1e6) / 1e6, nf: odNumFmt(r.fmt, cell.value, r.whole), style: odStyleOf(cell.st) };
  }));
  const change = rows.map((r) => { const ch = odChange(r, cols); return ch ? { v: xlText(ch.text), nf: "@", style: "chg-" + (ch.cls || "flat") } : { v: "", nf: "@", style: "none" }; });
  const fixed = rows.map((r) => [xlText(r.scope), xlText(r.name), xlText(r.formula), xlText(r.target)]);
  const sig = hashStr(JSON.stringify([ODV, kind, client.name, rows.map((r) => [r.key, r.scope, r.name, r.formula, r.target])]));
  const hashes = cols.map((c, j) => hashStr(JSON.stringify([head[4 + j], colCells[j].map((x) => [x.v, x.nf, x.style])])));
  const title = `${client.name}: Drishti summary, ${kind === "week" ? "weekly" : "monthly"}`;
  const empty = rows.length ? "" : `No ${kind === "week" ? "weeks" : "months"} with reports in Drishti yet. Upload one in Drishti, from OneDrive or by dropping the files, then click Update Drishti summary again.`;
  return { kind, title, stampText, head, rows, cols, colCells, change, fixed, sig, hashes, first: cols[0] ? cols[0].key : "", count: cols.length, empty };
}
// The summary's change column compares strictly the newest month (or week) with the one right before it.
// It stays blank when either of the two has no value or was not uploaded. (The History tab on screen
// keeps its own rule, newest available value against the one before.)
function odChange(row, cols) {
  const k = cols.length; if (k < 2 || row.isTxt) return null;
  if (!cols[k - 1].period || !cols[k - 2].period) return null;
  const a = row.cells[k - 2], b = row.cells[k - 1];
  if (!a || !b || !isNum(a.value) || !isNum(b.value)) return null;
  const d = b.value - a.value, abs = Math.abs(d), fmt = row.fmt;
  const txt = fmt === FMT.PERCENT ? (abs * 100).toFixed(abs * 100 >= 10 ? 0 : 1) + " pts" : fmt === FMT.DAYS ? Math.round(abs) + (Math.round(abs) === 1 ? " day" : " days") : Math.round(abs).toLocaleString();
  return { d, text: d === 0 ? "No change" : (d > 0 ? "▲ " : "▼ ") + txt, cls: d === 0 ? "flat" : deltaCls(d, row.rule || { dir: "info" }) };
}
// cells sharing a style, gathered into as few rectangles as possible so the formatting takes few calls
function styleRects(cells) {
  const byStyle = new Map();
  for (const c of cells) { if (!c.style || c.style === "none") continue; if (!byStyle.has(c.style)) byStyle.set(c.style, []); byStyle.get(c.style).push(c); }
  const out = [];
  for (const [style, list] of byStyle) {
    const rows = new Map(); for (const c of list) { if (!rows.has(c.r)) rows.set(c.r, []); rows.get(c.r).push(c.c); }
    const runs = [];
    for (const r of [...rows.keys()].sort((a, b) => a - b)) {
      const cs = rows.get(r).sort((a, b) => a - b); let a = cs[0], b = cs[0];
      for (let i = 1; i <= cs.length; i++) { if (i < cs.length && cs[i] === b + 1) { b = cs[i]; continue; } runs.push({ r1: r, r2: r, c1: a, c2: b }); if (i < cs.length) { a = cs[i]; b = cs[i]; } }
    }
    const merged = [];
    for (const run of runs) { const prev = merged.find((m) => m.c1 === run.c1 && m.c2 === run.c2 && m.r2 === run.r1 - 1); if (prev) prev.r2 = run.r2; else merged.push({ ...run }); }
    for (const m of merged) out.push({ style, ...m });
  }
  return out;
}
const OD_W = { A: 135, B: 190, C: 300, D: 80, month: 78, week: 96, change: 128 };
// Write one of Drishti's own tabs. A full write clears the tab and draws it again; otherwise only the
// months or weeks whose numbers changed, the change column and the "Updated by" line are rewritten.
async function odWriteTab(ctx, model, prev, sheetId, fresh) {
  const ws = ctx.base + "/worksheets/" + enc(sheetId);
  const R = (a) => `${ws}/range(address='${a}')`;
  const reqs = [];
  const nRows = 3 + model.rows.length, chgC = 4 + model.count, lastC = chgC;
  const periodW = model.kind === "week" ? OD_W.week : OD_W.month;
  const headerStyle = (c1, c2) => {
    const a = cellAddr(c1, 3, c2, 3);
    reqs.push({ method: "PATCH", url: R(a) + "/format/fill", body: { color: OD_INK } });
    reqs.push({ method: "PATCH", url: R(a) + "/format/font", body: { color: "#FFFFFF", bold: true } });
    reqs.push({ method: "PATCH", url: R(a) + "/format", body: { wrapText: true, verticalAlignment: "Center", horizontalAlignment: c1 >= 3 ? "Center" : "Left" } });
  };
  const bodyCells = (c1, c2) => {
    const cells = [];
    for (let c = c1; c <= c2; c++) model.rows.forEach((_, i) => { const x = c === chgC ? model.change[i] : model.colCells[c - 4][i]; cells.push({ r: 4 + i, c, style: x.style }); });
    return cells;
  };
  const paintBody = (c1, c2, reset) => {
    if (!model.rows.length) return;
    const a = cellAddr(c1, 4, c2, nRows);
    if (reset) { reqs.push({ method: "POST", url: R(a) + "/format/fill/clear", body: {} }); reqs.push({ method: "PATCH", url: R(a) + "/format/font", body: { color: OD_INK, bold: false } }); }
    reqs.push({ method: "PATCH", url: R(a) + "/format", body: { horizontalAlignment: "Center", verticalAlignment: "Center", wrapText: true } });
    for (const rc of styleRects(bodyCells(c1, c2))) {
      const s = OD_STYLE[rc.style]; if (!s) continue; const ra = cellAddr(rc.c1, rc.r1, rc.c2, rc.r2);
      if (s.fill) reqs.push({ method: "PATCH", url: R(ra) + "/format/fill", body: { color: s.fill } });
      if (s.font) reqs.push({ method: "PATCH", url: R(ra) + "/format/font", body: { color: s.font, bold: !!s.bold } });
    }
  };
  const block = (c1, c2, r1, r2) => {
    const values = [], formats = [];
    for (let r = r1; r <= r2; r++) {
      const vr = [], fr = [];
      for (let c = c1; c <= c2; c++) {
        let v = "", nf = "@";
        if (r === 1) v = c === 0 ? xlText(model.title) : "";
        else if (r === 2) v = c === 0 ? xlText(model.stampText) : "";
        else if (r === 3) v = xlText(model.head[c] || "");
        else { const i = r - 4; if (c < 4) v = model.fixed[i][c]; else if (c === chgC) { v = model.change[i].v; nf = model.change[i].nf; } else { const x = model.colCells[c - 4][i]; v = x.v; nf = x.nf; } }
        vr.push(v); fr.push(nf);
      }
      values.push(vr); formats.push(fr);
    }
    const a = cellAddr(c1, r1, c2, r2);
    // the format first, so a heading such as "Sep 2026" stays text and is not turned into a date
    reqs.push({ method: "PATCH", url: R(a), body: { numberFormat: formats } });
    reqs.push({ method: "PATCH", url: R(a), body: { values } });
  };
  let incremental = !fresh && !!prev && prev.sig === model.sig && prev.first === model.first && model.count >= prev.count && prev.count > 0 && Array.isArray(prev.hashes);
  if (incremental) incremental = await odLayoutIntact(ctx, R, model, prev);
  let columns = 0;
  if (!incremental) {
    if (!fresh) {
      const used = await ctx.get(`${ws}/usedRange(valuesOnly=false)?$select=address`);
      const a = used && used.address ? String(used.address).split("!").pop() : "";
      if (a) reqs.push({ method: "POST", url: R(a) + "/clear", body: { applyTo: "All" } });
    }
    block(0, lastC, 1, nRows);
    if (model.empty) { reqs.push({ method: "PATCH", url: R("A4:A4"), body: { numberFormat: [["@"]] } }); reqs.push({ method: "PATCH", url: R("A4:A4"), body: { values: [[xlText(model.empty)]] } }); reqs.push({ method: "PATCH", url: R("A4:A4") + "/format/font", body: { italic: true, color: OD_INK2 } }); }
    reqs.push({ method: "PATCH", url: R("A1:A1") + "/format/font", body: { bold: true, size: 14, color: OD_INK } });
    reqs.push({ method: "PATCH", url: R("A2:A2") + "/format/font", body: { size: 10, color: OD_INK2 } });
    headerStyle(0, lastC);
    reqs.push({ method: "PATCH", url: R("3:3") + "/format", body: { rowHeight: 32 } });
    [["A", OD_W.A], ["B", OD_W.B], ["C", OD_W.C], ["D", OD_W.D]].forEach(([l, w]) => reqs.push({ method: "PATCH", url: R(`${l}:${l}`) + "/format", body: { columnWidth: w } }));
    if (model.count) reqs.push({ method: "PATCH", url: R(`${colLetter(4)}:${colLetter(chgC - 1)}`) + "/format", body: { columnWidth: periodW } });
    reqs.push({ method: "PATCH", url: R(`${colLetter(chgC)}:${colLetter(chgC)}`) + "/format", body: { columnWidth: OD_W.change } });
    if (model.rows.length) {
      reqs.push({ method: "PATCH", url: R(cellAddr(0, 4, 3, nRows)) + "/format", body: { verticalAlignment: "Center", wrapText: true } });
      reqs.push({ method: "PATCH", url: R(cellAddr(0, 4, 0, nRows)) + "/format/font", body: { color: OD_INK2 } });
      reqs.push({ method: "PATCH", url: R(cellAddr(2, 4, 2, nRows)) + "/format/font", body: { color: OD_INK2, size: 10 } });
      reqs.push({ method: "PATCH", url: R(cellAddr(3, 4, 3, nRows)) + "/format", body: { horizontalAlignment: "Center" } });
      reqs.push({ method: "PATCH", url: R("A4:B4") + "/format/font", body: { bold: true } });
      paintBody(4, chgC, false);
    }
    columns = model.count;
  } else {
    const changed = []; for (let j = 0; j < model.count; j++) if (j >= prev.count || prev.hashes[j] !== model.hashes[j]) changed.push(j);
    const runs = []; for (const j of changed) { const last = runs[runs.length - 1]; if (last && last[1] === j - 1) last[1] = j; else runs.push([j, j]); }
    for (const [a, b] of runs) {
      const c1 = 4 + a, c2 = 4 + b;
      block(c1, c2, 3, nRows); headerStyle(c1, c2);
      reqs.push({ method: "PATCH", url: R(`${colLetter(c1)}:${colLetter(c2)}`) + "/format", body: { columnWidth: periodW } });
      paintBody(c1, c2, true);
    }
    block(chgC, chgC, 3, nRows); headerStyle(chgC, chgC);
    reqs.push({ method: "PATCH", url: R(`${colLetter(chgC)}:${colLetter(chgC)}`) + "/format", body: { columnWidth: OD_W.change } });
    paintBody(chgC, chgC, true);
    block(0, 0, 2, 2);
    columns = changed.length;
  }
  await ctx.run(reqs);
  return { state: { sig: model.sig, first: model.first, count: model.count, hashes: model.hashes }, columns, full: !incremental, calls: reqs.length };
}
// Before writing only some columns, check the tab still looks the way Drishti left it. Someone sorting
// or inserting rows or columns means a full rewrite instead, so nothing lands in the wrong place.
async function odLayoutIntact(ctx, R, model, prev) {
  try {
    const nRows = 3 + model.rows.length, chgPrev = 4 + prev.count;
    const h = await ctx.get(R(cellAddr(0, 3, chgPrev, 3)) + "?$select=values");
    const b = model.rows.length ? await ctx.get(R(cellAddr(1, 4, 1, nRows)) + "?$select=values") : { values: [] };
    const first = (s) => String(s == null ? "" : s).split("\n")[0].trim();
    const heads = (h && h.values && h.values[0]) || [];
    for (let c = 0; c < 4; c++) if (first(heads[c]) !== first(model.head[c])) return false;
    for (let j = 0; j < prev.count; j++) if (first(heads[4 + j]) !== first(model.head[4 + j])) return false;
    if (!/^Change vs/.test(first(heads[chgPrev]))) return false;
    const names = ((b && b.values) || []).map((r) => String(r[0] == null ? "" : r[0]).trim());
    return names.length === model.rows.length && names.every((nm, i) => nm === model.fixed[i][1].trim());
  } catch { return false; }
}
async function odWriteHowTo(ctx, sheetId, fresh) {
  const ws = ctx.base + "/worksheets/" + enc(sheetId), R = (a) => `${ws}/range(address='${a}')`, reqs = [];
  const lines = [
    ["How to read this summary", ""],
    ["", ""],
    ["Monthly tab", "One row for each KPI and SLA line, and one column for each month, oldest on the left and newest on the right."],
    ["Weekly tab", "The same for weeks. A week runs Monday to Friday and sits under the month its Monday is in, so Week 1 is the week of the month's first Monday."],
    ["Change column", "The last column on each tab compares the newest month (or week) with the one right before it. It is left blank when either of those two has no value or was not uploaded."],
    ["Colours", "Green is on target, amber is at risk, red is missed. A grey column is a month or week with no reports in Drishti."],
    ["Overall score", "Metrics on target ÷ metrics scored, the same score as Drishti's Overview."],
    ["Where the numbers come from", "The report files in this client's OneDrive folders, worked out by Drishti with its locked formulas. A value a supervisor typed over in Drishti shows as typed."],
    ["When it changes", "Drishti rewrites the Monthly and Weekly tabs each time someone clicks Update Drishti summary. Anything typed into those two tabs is replaced, so keep your own notes in a tab of your own. Drishti never changes other tabs or the report files."],
    ["Patient data", "None. This file holds KPI numbers only. No patient or claim rows are copied into it."],
  ];
  if (!fresh) {
    const used = await ctx.get(`${ws}/usedRange(valuesOnly=false)?$select=address`);
    const a = used && used.address ? String(used.address).split("!").pop() : "";
    if (a) reqs.push({ method: "POST", url: R(a) + "/clear", body: { applyTo: "All" } });
  }
  const a = cellAddr(0, 1, 1, lines.length);
  reqs.push({ method: "PATCH", url: R(a), body: { numberFormat: lines.map((l) => l.map(() => "@")) } });
  reqs.push({ method: "PATCH", url: R(a), body: { values: lines.map((l) => l.map(xlText)) } });
  reqs.push({ method: "PATCH", url: R("A1:A1") + "/format/font", body: { bold: true, size: 14, color: OD_INK } });
  reqs.push({ method: "PATCH", url: R(cellAddr(0, 3, 0, lines.length)) + "/format/font", body: { bold: true, color: OD_INK } });
  reqs.push({ method: "PATCH", url: R(cellAddr(0, 3, 1, lines.length)) + "/format", body: { wrapText: true, verticalAlignment: "Top" } });
  reqs.push({ method: "PATCH", url: R("A:A") + "/format", body: { columnWidth: 170 } });
  reqs.push({ method: "PATCH", url: R("B:B") + "/format", body: { columnWidth: 520 } });
  await ctx.run(reqs);
}
function starterWorkbook(clientName) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[`${clientName}: Drishti summary, monthly`]]), OD_TABS.month);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[`${clientName}: Drishti summary, weekly`]]), OD_TABS.week);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["How to read this summary"]]), OD_HOWTO);
  return new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" }));
}
// Find Drishti's summary file in the client's folder, or make it. A new file is only ever created, with
// "fail" if one already exists, so an existing summary is never replaced by an uploaded copy.
async function odEnsureSummary(od, clientName, known) {
  if (known && known.itemId) {
    try { const it = await msx.graph(odItem(od, known.itemId) + "?$select=id,name,webUrl,file,deleted"); if (it && it.file && !it.deleted) return { item: it, fresh: false }; }
    catch (e) { if (e.status !== 404) throw e; }
  }
  const name = summaryFileName(clientName), byPath = odItem(od, od.itemId) + ":/" + enc(name);
  try { const it = await msx.graph(byPath + "?$select=id,name,webUrl,file"); if (it && it.file) return { item: it, fresh: false }; }
  catch (e) { if (e.status !== 404) throw e; }
  try {
    const it = await msx.graph(byPath + ":/content?@microsoft.graph.conflictBehavior=fail", { method: "PUT", headers: { "Content-Type": XLSX_MIME }, body: starterWorkbook(clientName) });
    return { item: it, fresh: true };
  } catch (e) {
    if (e.status === 409) { const it = await msx.graph(byPath + "?$select=id,name,webUrl,file"); return { item: it, fresh: false }; } // someone made it a moment ago
    throw e;
  }
}
const odStampText = (who, d) => `Updated by ${who} on ${d.getDate()} ${MSHORT[d.getMonth()]} ${d.getFullYear()}, ${d.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" })}. Green is on target, amber is at risk, red is missed. Numbers come from the report files in this client's folders.`;
async function odUpdateSummary(client, me, onStep) {
  const od = client.onedrive; if (!od || !od.itemId) throw msErr("input", "Pick the folder for this client's summary file first.");
  const when = new Date(), who = personName(me) || me.username, stampText = odStampText(who, when);
  const prevSum = od.summary || {};
  if (onStep) onStep("Finding the summary file…");
  const { item, fresh } = await odEnsureSummary(od, client.name, prevSum);
  const base = odItem(od, item.id) + "/workbook";
  const partial = { itemId: item.id, name: item.name, webUrl: item.webUrl };
  const sess = await msx.graph(base + "/createSession", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ persistChanges: true }) });
  const sid = sess && sess.id;
  const ctx = { base, sid, get: (path) => msx.graph(path, { headers: { "workbook-session-id": sid } }), run: (reqs) => graphBatch(reqs, sid) };
  const prevTabs = !fresh && prevSum.itemId === item.id && prevSum.tabs ? prevSum.tabs : {};
  const tabs = {}, report = {};
  try {
    if (onStep) onStep("Opening the summary file…");
    const sheets = (((await ctx.get(base + "/worksheets?$select=id,name")) || {}).value) || [];
    const sheetFor = async (name) => {
      const s = sheets.find((x) => x.name === name); if (s) return { id: s.id, made: false };
      const made = await msx.graph(base + "/worksheets/add", { method: "POST", headers: { "Content-Type": "application/json", "workbook-session-id": sid }, body: JSON.stringify({ name }) });
      return { id: made.id, made: true };
    };
    for (const kind of ["month", "week"]) {
      const name = OD_TABS[kind]; if (onStep) onStep(`Writing the ${name} tab…`);
      const sh = await sheetFor(name);
      const res = await odWriteTab(ctx, odTabModel(client, kind, stampText), sh.made ? null : prevTabs[name], sh.id, fresh || sh.made);
      tabs[name] = res.state; report[name] = { columns: res.columns, full: res.full, calls: res.calls };
    }
    const how = await sheetFor(OD_HOWTO);
    if (fresh || how.made || (prevSum.howto || 0) < ODV || prevSum.itemId !== item.id) await odWriteHowTo(ctx, how.id, fresh || how.made);
  } catch (e) { e.partial = partial; throw e; }
  finally { try { await msx.graph(base + "/closeSession", { method: "POST", headers: { "Content-Type": "application/json", "workbook-session-id": sid }, body: "{}" }); } catch { /* the session lapses by itself */ } }
  return { ...partial, updatedBy: who, updatedAt: when.toISOString(), tabs, howto: ODV, report, fresh };
}
function whenText(iso) {
  const d = new Date(iso); if (!iso || isNaN(d)) return "";
  const t = d.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" }), now = new Date();
  const day = (x) => `${x.getFullYear()}-${x.getMonth()}-${x.getDate()}`;
  if (day(d) === day(now)) return "today " + t;
  if (day(d) === day(addDays(now, -1))) return "yesterday " + t;
  return `${d.getDate()} ${MSHORT[d.getMonth()]}${d.getFullYear() !== now.getFullYear() ? " " + d.getFullYear() : ""}, ${t}`;
}

/* ---- the OneDrive bar at the top of Reports for KPI: where the files come from, and the summary file ---- */
const OdCloud = () => (<svg width="22" height="16" viewBox="0 0 24 16" fill="none" stroke="#fff" strokeWidth="2" strokeLinejoin="round" aria-hidden="true"><path d="M6.5 15h12a4 4 0 0 0 .6-7.95A6 6 0 0 0 7.6 5.2 5 5 0 0 0 6.5 15z" /></svg>);
function OneDriveBar({ me, client, period, onSaveOd, picker }) {
  const supported = msSupported();
  const saved = (client && client.onedrive) || {};
  const od = saved.itemId ? saved : null;                       // the client's folder, where the summary goes
  const [acct, setAcct] = useState(null);
  const [busy, setBusy] = useState(null);       // "signin" while signing in or choosing a folder, "update" while writing the summary
  const [step, setStep] = useState("");
  const [err, setErr] = useState("");
  const [done, setDone] = useState("");
  const hint = odHint(me);
  useEffect(() => {
    msx.hint = hint || ""; if (!supported) return undefined;
    let live = true; msx.account(hint).then((a) => { if (live) setAcct(a || null); }).catch(() => {});
    const onAcct = (e) => { if (live) setAcct(e.detail || null); };
    window.addEventListener(OD_ACCT_EVENT, onAcct);
    return () => { live = false; window.removeEventListener(OD_ACCT_EVENT, onAcct); };
  }, [hint, supported]);
  useEffect(() => { setErr(""); setDone(""); }, [period && period.id, client && client.id]);
  const run = async (kind, fn) => {
    if (busy) return; setErr(""); setDone(""); setBusy(kind);
    try { await fn(); } catch (e) { setErr(msalMessage(e)); } finally { setBusy(null); setStep(""); }
  };
  const switchAccount = () => run("signin", async () => { const a = await msx.signIn(hint, true); await odCheckAccount(a); setAcct(a); odAnnounce(a); });
  // the folder the summary goes in: found from the first upload, or picked once
  const saveSummaryFolder = (root, trail, item) => {
    const cur = (client && client.onedrive) || {}, same = cur.itemId === item.id;
    const next = { ...cur, driveId: root.driveId, itemId: item.id, name: item.name || trail[trail.length - 1].name, webUrl: item.webUrl || "", path: trail.map((t) => t.name).join(" › "), trail, linkedBy: personName(me) || me.username, linkedAt: nowIso() };
    if (!same) delete next.summary;
    onSaveOd(next, "set the folder for the Drishti summary");
    return next;
  };
  const folderItem = async (root, trail) => { const last = trail[trail.length - 1]; return (await msx.graph(odItem({ driveId: root.driveId }, last.id) + "?$select=id,name,webUrl")) || last; };
  const askFolder = (then) => picker.current && picker.current.pick({ mode: "folder", why: `Open the folders down to ${client.name}'s folder, then click Use this folder. Drishti keeps the summary there and remembers it.`, onFolder: then });
  const writeSummary = async (folder) => {
    try {
      const res = await odUpdateSummary({ ...client, onedrive: folder }, me, setStep);
      onSaveOd({ ...folder, summary: { itemId: res.itemId, name: res.name, webUrl: res.webUrl, updatedBy: res.updatedBy, updatedAt: res.updatedAt, tabs: res.tabs, howto: res.howto } }, "updated Drishti summary");
      setDone(`Done. ${res.name} in ${folder.name} is up to date.`);
    } catch (e) {
      if (e.partial) onSaveOd({ ...folder, summary: { ...(folder.summary || {}), ...e.partial, tabs: {} } }); // the file was made: keep its link
      throw e;
    }
  };
  // Update Drishti summary: only when someone clicks. The first time, it asks which folder the summary goes in.
  const update = () => {
    if (od) { run("update", async () => { await odEnsure(hint, setStep); await writeSummary(od); }); return; }
    askFolder((root, trail) => run("update", async () => writeSummary(saveSummaryFolder(root, trail, await folderItem(root, trail)))));
  };
  const changeFolder = () => askFolder((root, trail) => run("signin", async () => { const f = saveSummaryFolder(root, trail, await folderItem(root, trail)); setDone(`The summary file will go in ${f.path}.`); }));
  if (!supported) return (
    <div className="od-bar od-off"><span className="od-ic sm"><OdCloud /></span>
      <span className="od-off-t">{!msConfigured() ? "OneDrive is not switched on for Drishti yet. Drop the files below as before." : `OneDrive works on the live Drishti site, ${MS_HOME.replace(/^https:\/\//, "")}. Here, drop the files below as before.`}</span>
    </div>
  );
  const acctBox = (<div className="od-acct">Microsoft account<br />{acct ? <><b>{acct.username}</b> · <button className="link" onClick={switchAccount} disabled={!!busy}>Switch</button></> : <><span className="od-muted">Not signed in on this computer</span> · <button className="link" onClick={switchAccount} disabled={!!busy}>Sign in</button></>}</div>);
  const src = period && period.source && period.source.from === "OneDrive" ? period.source : null;
  const sum = od && od.summary && od.summary.updatedAt ? od.summary : null;
  return (
    <div className="od-bar">
      <div className="od-top">
        <span className="od-ic"><OdCloud /></span>
        <div className="od-main">
          <div className="od-title">OneDrive</div>
          <div className="od-sub">Report files come from {odAccountShown() ? <>the OneDrive of <b>{odAccountShown()}</b></> : "OneDrive"}. Use <b>Upload from OneDrive</b> below.</div>
          {od && <div className="od-sub">Summary file goes in: {od.path || od.name} &nbsp;·&nbsp; <button className="link" onClick={changeFolder} disabled={!!busy}>Change</button>{od.webUrl && <> &nbsp;·&nbsp; <a className="link" href={od.webUrl} target="_blank" rel="noopener noreferrer">Open in OneDrive ↗</a></>}</div>}
        </div>
        {acctBox}
      </div>
      <div className="od-actions">
        <button className="btn" onClick={update} disabled={!!busy}>{busy === "update" ? "Updating…" : "Update Drishti summary"}</button>
        <span className="od-status">{step || (src ? <>Last uploaded from OneDrive {whenText(src.at)} by {src.by}{src.files ? `, ${src.files} file${src.files === 1 ? "" : "s"}` : ""}{src.folder ? <>, from {src.folder}</> : null}</> : <>Not uploaded from OneDrive yet for {period.label}</>)}<br />{sum ? <>Summary updated {whenText(sum.updatedAt)} by {sum.updatedBy}{sum.webUrl && <> · <a className="link" href={sum.webUrl} target="_blank" rel="noopener noreferrer">Open summary ↗</a></>}</> : <>Summary not written yet</>}</span>
      </div>
      {err && <div className="rc-err od-err" role="alert">{err}</div>}
      {done && !err && <div className="od-done">{done}</div>}
    </div>
  );
}
// The one OneDrive window behind every upload on a client's page: the big box, each report's own box, Pick again
// in the checklist, and choosing the summary's folder. It signs in (as the shared OneDrive account when
// config.js names one), opens My files or wherever this client's last upload came from, and hands the
// downloaded files back to whoever asked.
function OdPicker({ me, client, apiRef, onSaveOd }) {
  const [req, setReq] = useState(null);     // { mode: "many" | "one" | "folder", name, why, onFiles, onFolder }
  const [root, setRoot] = useState(null);
  const [work, setWork] = useState("");
  const [err, setErr] = useState("");
  const live = useRef(null);                // the request still on screen; a closed one is ignored when it finishes
  const hint = odHint(me);
  const saved = (client && client.onedrive) || {};
  useEffect(() => { msx.hint = hint || ""; msx.account(hint).catch(() => {}); }, [hint]); // gets Microsoft sign-in ready before the first click
  const open = async (r) => {
    live.current = r; setReq(r); setRoot(null); setErr(""); setWork("Opening OneDrive…");
    try {
      await odEnsure(hint, (t) => { if (live.current === r) setWork(t); });
      const rt = await odRoot(); if (live.current !== r) return;
      setWork(""); setRoot(rt);
    } catch (e) { if (live.current === r) { setWork(""); setErr(msalMessage(e)); } }
  };
  useEffect(() => { if (apiRef) apiRef.current = { pick: open }; });
  const close = () => { if (work && root) return; live.current = null; setReq(null); setRoot(null); setErr(""); setWork(""); };
  const startAt = (trail) => (Array.isArray(trail) && trail.length && trail[0].id === root.itemId ? trail : [{ id: root.itemId, name: root.name }]);
  const lastTrail = saved.lastFolder && saved.lastFolder.trail;
  const took = async (trail, items) => {
    setErr(""); setWork(`Reading 0 of ${items.length} file${items.length === 1 ? "" : "s"}…`);
    try {
      const drive = { driveId: root.driveId };
      const files = await odDownloadFiles(drive, items, (k, n) => setWork(`Reading ${k} of ${n} file${n === 1 ? "" : "s"}…`));
      // remember where they were, and (the first time) the client's folder, where the summary will go
      const cur = (client && client.onedrive) || {}, next = { ...cur, driveId: root.driveId, lastFolder: { trail, at: nowIso() } };
      if (!cur.itemId) {
        const ct = odClientTrail(trail);
        if (ct.length > 1) {
          const cf = ct[ct.length - 1];
          let webUrl = ""; try { const it = await msx.graph(odItem(drive, cf.id) + "?$select=id,webUrl"); webUrl = (it && it.webUrl) || ""; } catch { /* the link to open it is only a nicety */ }
          Object.assign(next, { itemId: cf.id, name: cf.name, webUrl, path: ct.map((t) => t.name).join(" › "), trail: ct, linkedBy: personName(me) || me.username, linkedAt: nowIso() });
        }
      }
      onSaveOd(next, null);
      const r = req, where = trail.slice(1).map((t) => t.name).join(" › ") || trail[0].name;
      live.current = null; setReq(null); setRoot(null); setWork("");
      if (r.onFiles) r.onFiles(files, { where, folderName: trail[trail.length - 1].name, count: items.length });
    } catch (e) { setWork(""); setErr(msalMessage(e)); }
  };
  const tookFolder = (trail) => { const r = req, rt = root; live.current = null; setReq(null); setRoot(null); if (r.onFolder) r.onFolder(rt, trail); };
  if (!req) return null;
  const title = req.mode === "folder" ? "Where should the summary file go?" : req.mode === "one" ? `Upload ${req.name || "a file"} from OneDrive` : "Upload files from OneDrive";
  const why = req.why || (req.mode === "one" ? "Open the folder with this report and tick it." : "Open the folder with the report files. Tick the ones you want, or take every file in the folder. Each one is matched to its report by its name.");
  if (!root) return (<Overlay onClick={close}>
    <div className="modal od-browser" onClick={(e) => e.stopPropagation()}>
      <h3>{title}</h3>
      {err ? <div className="rc-err" role="alert">{err}</div> : <p>{work || "Opening OneDrive…"}</p>}
      <div className="modal-actions"><button className="btn ghost" onClick={close}>{err ? "Close" : "Cancel"}</button></div>
    </div>
  </Overlay>);
  const start = req.mode === "folder" ? startAt(saved.trail || (lastTrail ? odClientTrail(lastTrail) : null)) : startAt(lastTrail);
  return <OdBrowser root={root} start={start} mode={req.mode} title={title} why={why} work={work} err={err} onClose={close} onFiles={took} onFolder={tookFolder} />;
}
// Browse the OneDrive. "many": open folders and tick report files, or take every file in the folder.
// "one": tick the one file for a single report. "folder": open a folder and use it.
function OdBrowser({ root, start, mode, title, why, work, err, onClose, onFiles, onFolder }) {
  const [trail, setTrail] = useState(start);
  const [view, setView] = useState({ id: null, folders: [], files: [], err: "" });
  const [sel, setSel] = useState(() => new Set());
  const here = trail[trail.length - 1];
  const drive = { driveId: root.driveId };
  useEffect(() => {
    let live = true; setSel(new Set());
    odChildren(drive, here.id).then((kids) => {
      if (!live) return;
      const byName = (a, b) => String(a.name).localeCompare(String(b.name), undefined, { numeric: true });
      setView({ id: here.id, folders: kids.filter((x) => x.folder).sort(byName), files: kids.filter((x) => isReportItem(x) && !(mode === "one" && /\.zip$/i.test(x.name))).sort(byName), err: "" });
    }).catch((e) => { if (live) setView({ id: here.id, folders: [], files: [], err: msalMessage(e) }); });
    return () => { live = false; };
  }, [here.id]);
  const loading = view.id !== here.id;
  const shown = loading ? { folders: [], files: [], err: "" } : view;
  const forFiles = mode !== "folder", one = mode === "one", busy = !!work;
  const toggle = (id) => setSel((s) => { if (one) return s.has(id) ? new Set() : new Set([id]); const n = new Set(s); if (n.has(id)) n.delete(id); else n.add(id); return n; });
  const allOn = shown.files.length > 0 && sel.size === shown.files.length;
  const picked = shown.files.filter((f) => sel.has(f.id));
  const take = one ? picked : picked.length ? picked : shown.files;
  const atTop = trail.length < 2;
  const label = busy ? work : loading ? "Upload" : one ? (picked.length ? "Upload this file" : "Tick a file") : !shown.files.length ? "No report files here" : picked.length ? `Upload ${picked.length} file${picked.length === 1 ? "" : "s"}` : shown.files.length === 1 ? "Upload this file" : `Upload all ${shown.files.length} files here`;
  return (<Overlay onClick={busy ? () => {} : onClose}>
    <div className={"modal od-browser" + (forFiles ? " wide" : "")} onClick={(e) => e.stopPropagation()}>
      <h3>{title}</h3>
      {why && <p>{why}</p>}
      <div className="od-crumbs">{trail.map((t, i) => <React.Fragment key={t.id}>{i > 0 && <span className="od-sep">›</span>}{i < trail.length - 1 ? <button className="link" onClick={() => !busy && setTrail(trail.slice(0, i + 1))}>{t.name}</button> : <b>{t.name}</b>}</React.Fragment>)}</div>
      <div className="bfp-list">
        {loading && <div className="od-muted">Opening the folder…</div>}
        {!loading && shown.folders.map((f) => <button key={f.id} className="bfp-item" disabled={busy} onClick={() => setTrail([...trail, { id: f.id, name: f.name }])}><span className="bfp-name">{f.name}</span><span className="bfp-count">{f.folder && f.folder.childCount != null ? `${f.folder.childCount} item${f.folder.childCount === 1 ? "" : "s"} ›` : "›"}</span></button>)}
        {forFiles && !one && !loading && shown.files.length > 1 && <label className="od-file od-all"><input type="checkbox" checked={allOn} disabled={busy} onChange={() => setSel(allOn ? new Set() : new Set(shown.files.map((f) => f.id)))} /><span className="od-file-n">Select all {shown.files.length} files</span></label>}
        {forFiles && !loading && shown.files.map((f) => <label key={f.id} className="od-file"><input type={one ? "radio" : "checkbox"} name="od-one" checked={sel.has(f.id)} disabled={busy} onChange={() => toggle(f.id)} /><span className="od-file-n">{f.name}</span><span className="od-file-d">{f.lastModifiedDateTime ? whenText(f.lastModifiedDateTime) : ""}</span></label>)}
        {!loading && !shown.err && !shown.folders.length && !(forFiles && shown.files.length) && <div className="od-muted">{forFiles ? "Nothing in here yet." : "No folders inside this one."}</div>}
      </div>
      {(shown.err || err) && <div className="rc-err" role="alert">{shown.err || err}</div>}
      <div className="modal-actions">
        <button className="btn ghost" onClick={onClose} disabled={busy}>Cancel</button>
        {forFiles
          ? <button className="btn primary" disabled={busy || loading || !take.length} onClick={() => onFiles(trail, take)}>{label}</button>
          : <button className="btn primary" disabled={loading || atTop} onClick={() => onFolder(trail)}>Use this folder</button>}
      </div>
    </div>
  </Overlay>);
}

/* ===================== bulk upload: drop zone + review ===================== */
const BULK_ACCEPT = /\.(xlsx|xls|csv)$/i;
const BULK_IN = /\.(xlsx|xls|csv|zip)$/i;
const isJunkName = (n) => /^(~\$|\.)/.test(n);
const folderOf = (f) => (f && f._folder != null ? f._folder : String((f && f.webkitRelativePath) || "").split("/").slice(0, -1).join("/"));
const folderLabel = (p) => (p ? p.split("/").filter(Boolean).slice(-1)[0] : "Files outside any folder");
// Open any zip in the drop and hand back the spreadsheets inside it, each tagged with the folder it sat in.
async function expandZips(files) {
  const out = [];
  for (const f of files) {
    if (!/\.zip$/i.test(f.name)) { if (BULK_ACCEPT.test(f.name) && !isJunkName(f.name)) out.push(f); continue; }
    try {
      const z = XLSX.CFB.read(new Uint8Array(await f.arrayBuffer()), { type: "array" });
      z.FullPaths.forEach((p, i) => {
        const entry = z.FileIndex[i]; const parts = p.replace(/^Root Entry\//, "").split("/"); const base = parts[parts.length - 1];
        if (!base || !BULK_ACCEPT.test(base) || isJunkName(base) || parts.includes("__MACOSX") || !entry || !entry.content || !entry.content.length) return;
        const file = new File([new Uint8Array(entry.content)], base); file._folder = parts.slice(0, -1).join("/");
        out.push(file);
      });
    } catch { const bad = new File([], f.name); bad._unreadable = true; bad._folder = ""; out.push(bad); }
  }
  return out;
}
// Pull every spreadsheet out of a drop, including whole folders dragged in.
async function filesFromDrop(dt) {
  const items = dt && dt.items ? Array.from(dt.items) : [];
  const entries = items.map((it) => (it.webkitGetAsEntry ? it.webkitGetAsEntry() : null)).filter(Boolean);
  if (!entries.length) return Array.from((dt && dt.files) || []).filter((f) => BULK_IN.test(f.name));
  const out = [];
  const walk = async (entry, depth) => {
    if (depth > 4 || out.length > 200) return;
    if (entry.isFile) { const f = await new Promise((res) => entry.file(res, () => res(null))); if (f && BULK_IN.test(f.name)) { f._folder = String(entry.fullPath || "").replace(/^\//, "").split("/").slice(0, -1).join("/"); out.push(f); } return; }
    if (entry.isDirectory) {
      const reader = entry.createReader(); let batch;
      do { batch = await new Promise((res) => reader.readEntries(res, () => res([]))); for (const e of batch) await walk(e, depth + 1); } while (batch && batch.length);
    }
  };
  for (const e of entries) await walk(e, 0);
  return out;
}

function BulkDrop({ count, total, busy, onFiles, onCloud }) {
  const [drag, setDrag] = useState(false);
  const fileRef = useRef(null), dirRef = useRef(null);
  // OneDrive switched on: the period's files are picked there, never from the computer
  if (onCloud) return (
    <div className="bulkzone">
      <div className="bz-icon">⇪</div>
      <div className="bz-main">
        <div className="bz-title">{busy ? "Reading your files…" : "Upload this period's files from OneDrive, all at once"}</div>
        <div className="bz-sub">Pick the files in OneDrive. Every file is matched to its slot below and shown to you before anything is calculated. A zip file works too.</div>
        <div className="row" style={{ marginTop: 11, gap: 8 }}>
          <button className="btn sm primary" onClick={onCloud} disabled={busy}>Upload from OneDrive</button>
          <span className="bz-count">{count} of {total} slots filled</span>
        </div>
      </div>
    </div>
  );
  const pick = (e) => { const fs = Array.from(e.target.files || []).filter((f) => BULK_IN.test(f.name)); e.target.value = ""; if (fs.length) onFiles(fs); };
  const drop = async (e) => { e.preventDefault(); setDrag(false); const fs = await filesFromDrop(e.dataTransfer); if (fs.length) onFiles(fs); };
  return (
    <div className={"bulkzone" + (drag ? " on" : "")}
      onDragOver={(e) => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)} onDrop={drop}>
      <div className="bz-icon">⇪</div>
      <div className="bz-main">
        <div className="bz-title">{busy ? "Reading your files…" : "Drop this period's files here, all at once"}</div>
        <div className="bz-sub">Every file is matched to its slot below and shown to you before anything is calculated. A whole folder or a zip file works too.</div>
        <div className="row" style={{ marginTop: 11, gap: 8 }}>
          <button className="btn sm primary" onClick={() => fileRef.current && fileRef.current.click()} disabled={busy}>Choose files</button>
          <button className="btn sm" onClick={() => dirRef.current && dirRef.current.click()} disabled={busy}>Choose a folder</button>
          <span className="bz-count">{count} of {total} slots filled</span>
        </div>
      </div>
      <input ref={fileRef} type="file" multiple accept=".xlsx,.xls,.csv,.zip" style={{ display: "none" }} onChange={pick} />
      <input ref={dirRef} type="file" webkitdirectory="" directory="" multiple style={{ display: "none" }} onChange={pick} />
    </div>
  );
}

function BulkFolderPick({ folders, onPick, onCancel }) {
  return (
    <Overlay onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h3>Which folder is for this client?</h3>
        <p>What you dropped has more than one folder of reports. Pick the one that belongs to this client, so two sets of files are not mixed together.</p>
        <div className="bfp-list">
          {folders.map((f) => <button key={f.path} className="bfp-item" onClick={() => onPick(f.path)}><span className="bfp-name">{folderLabel(f.path)}</span><span className="bfp-count">{f.count} file{f.count === 1 ? "" : "s"}</span></button>)}
        </div>
        <div className="modal-actions"><button className="btn ghost" onClick={onCancel}>Cancel</button><button className="btn" onClick={() => onPick(null)}>Use every folder</button></div>
      </div>
    </Overlay>
  );
}
const BULK_STATUS = {
  certain: { label: "Matched", cls: "ok" },
  check: { label: "Please confirm", cls: "warn" },
  conflict: { label: "Clashes", cls: "miss" },
  unknown: { label: "Not recognised", cls: "info" },
  outofscope: { label: "Not in this client", cls: "info" },
};
function BulkReview({ rows, allowedKeys, onCancel, onApply, title, desc, onOtherFolder }) {
  const [picks, setPicks] = useState(() => rows.map((r) => (r.status === "conflict" ? "" : r.slot || "")));
  const set = (i, v) => setPicks((p) => p.map((x, j) => (j === i ? v : x)));
  const used = {}; picks.forEach((p, i) => { if (p) (used[p] = used[p] || []).push(i); });
  const clash = Object.entries(used).filter(([, list]) => list.length > 1).map(([k]) => k);
  const chosen = picks.filter(Boolean).length;
  const needs = rows.filter((r, i) => r.status !== "certain" && !picks[i]).length;
  const slotOptions = (allowedKeys || []).map((k) => ({ k, name: REPORT_BY_KEY[k] ? REPORT_BY_KEY[k].name : k }));
  return (
    <Overlay onClick={onCancel}>
      <div className="modal wide" onClick={(e) => e.stopPropagation()}>
        <h3>{title || "Check these before they are applied"}</h3>
        <p>{desc || "Nothing is calculated until you press apply. Change any row that is not right, or leave it blank to skip the file."}</p>
        {onOtherFolder && <div className="od-review-src">Not the right files? <button className="link" onClick={onOtherFolder}>Pick again in OneDrive</button></div>}
        {rows.length === 0 && <div className="rc-err" style={{ marginBottom: 10 }}>No Excel or CSV files were found in what you {onOtherFolder ? "picked" : "dropped"}.</div>}
        <div className="bulktable" style={rows.length === 0 ? { display: "none" } : undefined}>
          {rows.map((r, i) => {
            const st = r.outOfScope ? BULK_STATUS.outofscope : (BULK_STATUS[r.status] || BULK_STATUS.unknown);
            const dup = picks[i] && clash.includes(picks[i]);
            return (
              <div key={i} className={"bulkrow" + (dup ? " dup" : "")}>
                <div className="br-file">
                  <div className="br-name">{r.name}</div>
                  <div className="br-reason">{dup ? "Another file is already going to this slot" : r.outOfScope ? `Looks like the ${REPORT_BY_KEY[r.outOfScope]?.name || "report"}, but ${SCOPE_BY_KEY[RULE_BY_KEY[r.outOfScope]?.scope]?.name || "that scope"} is not switched on for this client` : r.reason}</div>
                </div>
                <span className={"chip " + st.cls}>{dup ? "Clashes" : st.label}</span>
                <select className="br-sel" value={picks[i]} onChange={(e) => set(i, e.target.value)}>
                  <option value="">Skip this file</option>
                  {slotOptions.map((o) => <option key={o.k} value={o.k}>{o.name}</option>)}
                </select>
              </div>
            );
          })}
        </div>
        <div className="bulkfoot">
          {chosen} file{chosen === 1 ? "" : "s"} will be applied{rows.length - chosen > 0 ? `, ${rows.length - chosen} skipped` : ""}.
          {needs > 0 ? ` ${needs} still need${needs === 1 ? "s" : ""} a slot chosen.` : ""}
          {clash.length > 0 ? " Two files are pointing at the same slot." : ""}
        </div>
        <div className="modal-actions">
          <button className="btn ghost" onClick={onCancel}>Cancel</button>
          <button className="btn primary" disabled={!chosen || clash.length > 0} onClick={() => onApply(picks)}>Apply {chosen || ""} file{chosen === 1 ? "" : "s"}</button>
        </div>
      </div>
    </Overlay>
  );
}

function ColumnPicker({ pending, onClose, onResolve }) {
  const cols = pending.derived.columns || [];
  const missing = pending.derived.missing || [], amb = pending.derived.ambiguous || [];
  const [map, setMap] = useState({});
  const set = (id, v) => setMap((m) => ({ ...m, [id]: v }));
  const nameOf = (i) => (cols.find((c) => c.index === i) || {}).name;
  // every item that was asked about goes back, blank included, so nothing is guessed afterwards
  const submit = () => { const out = {}; for (const f of [...missing, ...amb]) out[f.id] = map[f.id] ?? ""; onResolve(out); };
  const unanswered = amb.filter((f) => (map[f.id] ?? "") === "").length;
  return (<Modal onClose={onClose} title={`${REPORT_BY_KEY[pending.key].name}: which columns?`} desc={`Drishti could not be sure about ${missing.length + amb.length === 1 ? "one column" : "some columns"} in ${pending.derived.fileName || "this file"}. Tell it which column to read, and the numbers will be taken from there.`}>
    {amb.map((f) => (
      <div className="field" key={f.id}><label>{f.label}</label>
        <div className="cp-why">More than one column could be this: {f.options.map(nameOf).filter(Boolean).join(", ")}. Which one should be used?</div>
        <select value={map[f.id] ?? ""} onChange={(e) => set(f.id, e.target.value)}>
          <option value="">Select the column</option>
          <optgroup label="Could be">{f.options.map((i) => <option key={i} value={i}>{nameOf(i)}</option>)}</optgroup>
          <optgroup label="Other columns in this file">{cols.filter((c) => !f.options.includes(c.index)).map((c) => <option key={c.index} value={c.index}>{c.name}</option>)}</optgroup>
        </select>
      </div>
    ))}
    {missing.map((f) => (
      <div className="field" key={f.id}><label>{f.label}</label>
        <div className="cp-why">No column with a matching heading and the right kind of data was found.</div>
        <select value={map[f.id] ?? ""} onChange={(e) => set(f.id, e.target.value)}>
          <option value="">Select the column</option>
          {cols.map((c) => <option key={c.index} value={c.index}>{c.name}</option>)}
        </select>
      </div>
    ))}
    <div className="hint" style={{ fontSize: 11.5, color: "var(--ink-3)" }}>Leave one blank only if the file does not have it. That value will stay empty and can be typed in. Your choices are remembered for this client and this report.</div>
    <Actions onClose={onClose} onOk={submit} okLabel={unanswered ? "Use these columns (some left blank)" : "Use these columns"} />
  </Modal>);
}
function Confetti({ trigger }) {
  const [on, setOn] = useState(false);
  useEffect(() => { if (trigger > 0) { setOn(true); const t = setTimeout(() => setOn(false), 1900); return () => clearTimeout(t); } }, [trigger]);
  if (!on) return null;
  const colors = ["#3ECF7C", "#2CC9B0", "#38C8E8", "#0b7a4f", "#c07d10", "#f6c445"];
  return (<div className="confetti-wrap" aria-hidden="true">{Array.from({ length: 90 }, (_, i) => {
    const size = 6 + Math.random() * 6;
    return <span key={i} className="confetti-pc" style={{ left: Math.random() * 100 + "%", background: colors[i % colors.length], width: size, height: size * 0.6, animationDelay: Math.random() * 0.3 + "s", animationDuration: (1.1 + Math.random() * 0.8) + "s" }} />;
  })}</div>);
}

function ClientWorkspace({ me, can, clientId, onBack, onScopesChanged }) {
  const [client, setClient] = useState(null);
  const [pIdx, setPIdx] = useState(-1);
  const [scopeOpen, setScopeOpen] = useState(null);
  const [viewFilter, setViewFilter] = useState("both");
  const [addScope, setAddScope] = useState(false);
  const [editScopes, setEditScopes] = useState(false);
  const [reps, setReps] = useState({});
  const [numDays, setNumDays] = useState(30);
  const [pending, setPending] = useState(null);
  const [bulk, setBulk] = useState(null);
  const [bulkBusy, setBulkBusy] = useState(false);
  const [bulkPick, setBulkPick] = useState(null);
  const colQueueRef = useRef([]);
  const odPickRef = useRef(null);
  const [joinPending, setJoinPending] = useState(null);
  const [catEdit, setCatEdit] = useState(false);
  const [celebrate, setCelebrate] = useState(0);
  const [tab, setTab] = useState("period");          // "period" to work on one period, "history" to see them side by side
  const [loadingPeriod, setLoadingPeriod] = useState(false);
  const clientRef = useRef(null); useEffect(() => { clientRef.current = client; }, [client]);
  const repsRef = useRef({}); useEffect(() => { repsRef.current = reps; }, [reps]);
  const numDaysRef = useRef(30); useEffect(() => { numDaysRef.current = numDays; }, [numDays]);
  const persistRef = useRef(null);
  useEffect(() => { (async () => { const c = await sGet(clientKey(clientId)); if (c) c.periods = sortPeriods(c.periods); setClient(c); setPIdx(c?.periods?.length ? 0 : -1); })(); }, [clientId]);
  const reload = async () => { const c = await sGet(clientKey(clientId)); if (c) c.periods = sortPeriods(c.periods); setClient(c); return c; };
  const period = pIdx >= 0 ? client?.periods?.[pIdx] : null;
  const activeScopes = (client?.scopes || []).map((k) => SCOPE_BY_KEY[k]).filter(Boolean);
  // every upload slot this client's scopes can use, KPI slots plus the SLA extras
  const bulkSlots = useMemo(() => {
    const keys = [];
    for (const sc of activeScopes) { for (const k of scopeReportKeys(sc.key, client?.scopes)) if (!keys.includes(k)) keys.push(k); }
    for (const sc of activeScopes) { for (const k of SCOPE_SLA_REPORTS[sc.key] || []) if (!keys.includes(k)) keys.push(k); }
    return keys;
  }, [client?.scopes?.join(",")]);
  useEffect(() => { setReps(period?.reports || {}); setNumDays(period?.numDays || (period?.kind === "week" ? 7 : 30)); }, [pIdx, client?.id, period?.id, period?._stub]);
  // an older period arrives as a short copy; fetch the whole of it before anyone works on it
  useEffect(() => {
    const p = pIdx >= 0 ? clientRef.current?.periods?.[pIdx] : null;
    if (!p || !p._stub || !client) return;
    let live = true; setLoadingPeriod(true);
    (async () => {
      const full = await loadFullPeriod(client.id, p.id);
      if (!live) return; setLoadingPeriod(false); if (!full) return;
      const c = clientRef.current; const next = { ...c, periods: (c.periods || []).map((x) => (x.id === full.id ? full : x)) };
      clientRef.current = next; setClient(next);
    })();
    return () => { live = false; };
  }, [pIdx, client?.id, period?._stub]);

  // picking a month, week or day opens it if it already exists, otherwise starts it
  const selectPeriod = async ({ kind, from, to, label, numDays }) => {
    const c = clientRef.current || (await reload());
    setTab("period");
    const idx = (c.periods || []).findIndex((p) => p.range?.from === from && p.range?.to === to);
    if (idx >= 0) { setPIdx(idx); return; }
    const p = { id: uid(), label, kind, range: { from, to }, numDays, createdAt: nowIso(), createdBy: me.username, scopes: {}, reports: {} };
    const periods = sortPeriods([p, ...(c.periods || [])]); const next = { ...c, periods };
    clientRef.current = next; setClient(next); setPIdx(periods.findIndex((x) => x.id === p.id)); sSet(clientKey(c.id), next).catch(() => {});
    Promise.resolve(appendAudit({ user: me.username, role: me.role, action: "created period", client: c.name, period: label })).catch(() => {});
  };
  const saveScopeData = async (scopeKey, data, auditLabel, method, file) => {
    const c = clientRef.current; if (!c) return; const periods = [...c.periods]; const p = { ...periods[pIdx] };
    p.scopes = { ...p.scopes, [scopeKey]: data }; periods[pIdx] = p;
    const next = { ...c, periods }; clientRef.current = next; setClient(next); schedulePersist();
    Promise.resolve(appendAudit({ user: me.username, role: me.role, action: auditLabel, client: c.name, period: p.label, method: method || null, file: file || null })).catch(() => {});
  };
  const addScopeToClient = async (key) => {
    const c = clientRef.current; if (!c || (c.scopes || []).includes(key)) return setAddScope(false);
    const next = { ...c, scopes: [...(c.scopes || []), key] }; clientRef.current = next; setClient(next); setAddScope(false); sSet(clientKey(c.id), next).catch(() => {});
    Promise.resolve(appendAudit({ user: me.username, role: me.role, action: `added scope ${SCOPE_BY_KEY[key].short}`, client: c.name })).catch(() => {}); onScopesChanged && onScopesChanged();
  };
  const saveScopesList = async (scopes) => {
    const c = clientRef.current; if (!c) return; const next = { ...c, scopes }; clientRef.current = next; setClient(next); setEditScopes(false); sSet(clientKey(c.id), next).catch(() => {});
    Promise.resolve(appendAudit({ user: me.username, role: me.role, action: "updated scopes", client: c.name })).catch(() => {}); onScopesChanged && onScopesChanged();
  };
  const schedulePersist = () => { if (persistRef.current) clearTimeout(persistRef.current); persistRef.current = setTimeout(() => { const cl = clientRef.current; if (cl) sSet(clientKey(cl.id), cl).catch(() => {}); }, 250); };
  const recompute = (nextReps, days) => {
    const c = clientRef.current; if (!c || pIdx < 0) return { auto: {}, next: c };
    const periods = [...c.periods]; const p = { ...periods[pIdx] };
    const auto = autoCompute(nextReps, days, c.joinKeys);
    const scopes = JSON.parse(JSON.stringify(p.scopes || {}));
    const joinNeeds = [];
    for (const [fk, info] of Object.entries(auto)) {
      if (info.needsJoin) { joinNeeds.push(info.needsJoin); continue; }
      if (info.value === undefined && !info.text) continue;
      const [sk, ck] = fk.split("."); scopes[sk] = scopes[sk] || { crit: {} }; scopes[sk].crit = scopes[sk].crit || {};
      const prevCell = scopes[sk].crit[ck];
      if (prevCell && prevCell.method === "manual" && prevCell.value != null && prevCell.value !== "") {
        // the supervisor typed this value: keep it, and keep the calculated one beside it so it can be restored
        scopes[sk].crit[ck] = { ...prevCell, calc: { value: info.value, src: info.src, inputs: info.inputs || null } };
      } else scopes[sk].crit[ck] = { value: info.value, method: "upload", src: info.src, inputs: info.inputs || null, by: me.username, at: nowIso() };
    }
    if (joinNeeds.length && !clientRef.current?.joinKeys?.[joinNeeds[0].pair]) setTimeout(() => setJoinPending(joinNeeds[0]), 0);
    if (nextReps.ar?.aging) { scopes.ar = scopes.ar || { crit: {} }; scopes.ar.aging = { rows: Object.entries(nextReps.ar.aging).map(([bucket, amount]) => ({ bucket, amount })) }; }
    p.reports = nextReps; p.numDays = days; p.scopes = scopes; periods[pIdx] = p;
    const next = { ...c, periods }; clientRef.current = next; setClient(next); schedulePersist();
    return { auto, next };
  };
  const applyAndCelebrate = (key, derived) => {
    const stripped = { ...derived }; delete stripped.columns; delete stripped.missing; delete stripped.ambiguous;
    const nextReps = { ...repsRef.current, [key]: stripped }; repsRef.current = nextReps; setReps(nextReps);
    const { auto, next } = recompute(nextReps, numDaysRef.current);
    Promise.resolve(appendAudit({ user: me.username, role: me.role, action: `uploaded ${REPORT_BY_KEY[key].name}`, client: next?.name, period: period?.label, method: "upload", file: stripped.fileName })).catch(() => {});
    let best = false;
    for (const [fk, info] of Object.entries(auto)) { const [sk, ck] = fk.split("."); const crit = critOf(sk, ck); if (!crit) continue; const st = statusFor(info.value, effectiveRule(next, sk, crit)); if (st === "great" || st === "ok") { best = true; break; } }
    if (best) setCelebrate((x) => x + 1);
  };
  // apply several reports in one pass so the KPIs are recalculated once, not 22 times
  const applyMany = (list) => {
    if (!list.length) return;
    const nextReps = { ...repsRef.current };
    for (const { key, derived } of list) { const s = { ...derived }; delete s.columns; delete s.missing; delete s.ambiguous; nextReps[key] = s; }
    repsRef.current = nextReps; setReps(nextReps);
    const { auto, next } = recompute(nextReps, numDaysRef.current);
    Promise.resolve(appendAudit({ user: me.username, role: me.role, action: `bulk upload of ${list.length} report${list.length > 1 ? "s" : ""}`, client: next?.name, period: period?.label, method: "bulk upload", file: list.map((x) => x.derived.fileName).filter(Boolean).join(", ").slice(0, 400) })).catch(() => {});
    let best = false;
    for (const [fk, info] of Object.entries(auto)) { const [sk, ck] = fk.split("."); const crit = critOf(sk, ck); if (!crit) continue; const st = statusFor(info.value, effectiveRule(next, sk, crit)); if (st === "great" || st === "ok") { best = true; break; } }
    if (best) setCelebrate((x) => x + 1);
  };
  // read every dropped file, work out where each one belongs, then hand the list back for checking
  // read every dropped file, work out where each one belongs, then hand the list back for checking
  const readAndMatch = async (files, meta) => {
    setBulkBusy(true);
    const parsed = [];
    for (const f of files) {
      if (f._unreadable) { parsed.push({ name: f.name, headers: [], rep: null, unreadable: true }); continue; }
      try { const rep = await readReport(f); parsed.push({ name: f.name, headers: rep.headers, rep }); }
      catch { parsed.push({ name: f.name, headers: [], rep: null, unreadable: true }); }
    }
    const allowed = bulkSlots;
    const rows = matchFiles(parsed.map((p) => ({ name: p.name, headers: p.headers })), allowed)
      .map((r, i) => (parsed[i].unreadable
        ? { ...r, slot: null, status: "unknown", reason: "This file could not be read", rep: null }
        : { ...r, rep: parsed[i].rep }));
    setBulkBusy(false);
    setBulk({ rows, meta: meta || null });
  };
  // files from a drop, a folder, a zip, or a OneDrive fetch (meta says which folder they came from)
  const onBulkFiles = async (files, meta) => {
    setBulkBusy(true);
    const flat = await expandZips(files);
    setBulkBusy(false);
    if (!flat.length) { setBulk({ rows: [], meta: meta || null }); return; }
    const counts = {}; for (const f of flat) { const p = folderOf(f); counts[p] = (counts[p] || 0) + 1; }
    const folders = Object.entries(counts).map(([path, count]) => ({ path, count }));
    if (folders.length > 1) { setBulkPick({ folders, files: flat, meta: meta || null }); return; }
    await readAndMatch(flat, meta);
  };
  // OneDrive switched on: every upload on this page is picked in OneDrive, through the one window below
  const cloud = msSupported();
  const fromOd = (info) => ({ title: `From OneDrive: ${info.where}`, desc: `${info.count} file${info.count === 1 ? "" : "s"} from ${info.folderName}. Check where each one goes, then apply. Nothing is calculated until you do.`, source: { folder: info.where } });
  const pickBulk = () => { if (odPickRef.current) odPickRef.current.pick({ mode: "many", onFiles: (files, info) => onBulkFiles(files, fromOd(info)) }); };
  const pickOne = (def, then) => { if (odPickRef.current) odPickRef.current.pick({ mode: "one", name: def.name, onFiles: (files, info) => then(files[0], { folder: info.where }) }); };
  const pickBulkFolder = async (path) => { const all = bulkPick?.files || [], meta = bulkPick?.meta || null; setBulkPick(null); await readAndMatch(path == null ? all : all.filter((f) => folderOf(f) === path), meta); };
  // the client's OneDrive folder and summary bookkeeping are saved with its other settings
  const saveOd = (od, action) => {
    const c = clientRef.current; if (!c) return;
    const next = { ...c, onedrive: od }; clientRef.current = next; setClient(next); sSet(clientKey(c.id), next).catch(() => {});
    if (action) Promise.resolve(appendAudit({ user: me.username, role: me.role, action, client: c.name, period: action === "updated Drishti summary" ? null : period?.label || null, file: od?.path || od?.name || null })).catch(() => {});
  };
  // a period filled from OneDrive remembers the folder, who uploaded it and when
  const markFetched = (src, count) => {
    const c = clientRef.current; if (!c || pIdx < 0 || !c.periods?.[pIdx]) return;
    const by = personName(me) || me.username, at = nowIso();
    const periods = [...c.periods]; const p = { ...periods[pIdx], source: { from: "OneDrive", folder: src.folder, by, at, files: count } }; periods[pIdx] = p;
    const next = { ...c, periods, onedrive: { ...(c.onedrive || {}), lastFetch: { by, at, period: p.label, folder: src.folder, files: count } } };
    clientRef.current = next; setClient(next); schedulePersist();
    Promise.resolve(appendAudit({ user: me.username, role: me.role, action: "uploaded from OneDrive", client: c.name, period: p.label, method: "OneDrive", file: src.folder })).catch(() => {});
  };
  const applyBulk = (picks) => {
    const meta = bulk?.meta || null;
    const rows = bulk?.rows || []; const ready = [], needCols = [];
    rows.forEach((r, i) => {
      const key = picks[i]; if (!key || !r.rep) return;
      const base = savedCols(key, r.rep);
      const derived = deriveReport(key, r.rep, base);
      if (needsAsking(derived)) needCols.push({ key, rep: r.rep, derived, base });
      else ready.push({ key, derived });
    });
    setBulk(null);
    applyMany(ready);
    if (meta?.source) markFetched(meta.source, ready.length + needCols.length);
    if (needCols.length) { colQueueRef.current = needCols.slice(1); setTimeout(() => setPending(needCols[0]), 0); }
  };
  const nextInColQueue = () => { const q = colQueueRef.current; if (!q.length) { setPending(null); return; } const [nx, ...rest] = q; colQueueRef.current = rest; setPending(nx); };
  // columns the supervisor chose before for this client and report, matched by heading so a moved column still works
  const savedCols = (key, rep) => {
    const saved = clientRef.current?.colMap?.[key]; if (!saved) return undefined;
    const ov = {}; for (const [fid, name] of Object.entries(saved)) { const i = rep.headers.findIndex((h) => String(h).trim().toLowerCase() === String(name).trim().toLowerCase()); if (i >= 0) ov[fid] = i; }
    return Object.keys(ov).length ? ov : undefined;
  };
  const needsAsking = (d) => (d.missing && d.missing.length) || (d.ambiguous && d.ambiguous.length);
  const rememberCols = (key, rep, override) => {
    const c = clientRef.current; if (!c) return;
    const picked = {}; for (const [fid, v] of Object.entries(override || {})) if (v !== "" && v != null && rep.headers[Number(v)] != null) picked[fid] = rep.headers[Number(v)];
    if (!Object.keys(picked).length) return;
    const next = { ...c, colMap: { ...(c.colMap || {}), [key]: { ...((c.colMap || {})[key] || {}), ...picked } } };
    clientRef.current = next; setClient(next); sSet(clientKey(c.id), next).catch(() => {});
  };
  const forgetCols = (key) => {
    const c = clientRef.current; if (!c?.colMap?.[key]) return;
    const colMap = { ...c.colMap }; delete colMap[key];
    const next = { ...c, colMap }; clientRef.current = next; setClient(next); sSet(clientKey(c.id), next).catch(() => {});
  };
  const onReport = (key, rep, src) => {
    const base = savedCols(key, rep);
    const derived = deriveReport(key, rep, base);
    if (needsAsking(derived)) setPending({ key, rep, derived, base });
    else applyAndCelebrate(key, derived);
    if (src) markFetched(src, 1);
  };
  const resolveColumns = (override) => {
    if (!pending) return; const full = { ...(pending.base || {}), ...override };
    const derived = deriveReport(pending.key, pending.rep, full);
    rememberCols(pending.key, pending.rep, override);
    const key = pending.key; nextInColQueue(); applyAndCelebrate(key, derived);
  };
  const onDays = (d) => { setNumDays(d); numDaysRef.current = d; recompute(repsRef.current, d); };
  const clearReport = (key) => { const nextReps = { ...repsRef.current }; delete nextReps[key]; repsRef.current = nextReps; setReps(nextReps); recompute(nextReps, numDaysRef.current); };
  const resolveJoin = (fields) => {
    const c = clientRef.current; if (!c || !joinPending) return;
    const next = { ...c, joinKeys: { ...(c.joinKeys || {}), [joinPending.pair]: { fields } } };
    clientRef.current = next; setClient(next); sSet(clientKey(c.id), next).catch(() => {});
    setJoinPending(null); recompute(repsRef.current, numDaysRef.current);
  };
  const saveCatMapping = (map) => {
    const c = clientRef.current; if (!c) return;
    const next = { ...c, statusCategories: { ...(c.statusCategories || {}), ...map } };
    clientRef.current = next; setClient(next); sSet(clientKey(c.id), next).catch(() => {}); setCatEdit(false);
  };

  if (!client) return <div className="wrap"><div className="card" style={{ color: "var(--ink-3)" }}>Loading…</div></div>;
  return (
    <div className="wrap">
      <button className="link" style={{ marginBottom: 12 }} onClick={onBack}>← All clients</button>
      <div className="pagehead">
        <div>
          <h1>{client.name}</h1>
          <div className="sub" style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginTop: 6 }}>
            {REMOTE && client.ownerName && <span className="ownertag">Maintained by {client.ownerId === me.id ? "you" : client.ownerName}</span>}
            {can("scopes:edit") && <button className="link" onClick={() => setEditScopes(true)}>edit scopes</button>}
          </div>
        </div>
        <ViewSwitch tab={tab} onTab={setTab} />
      </div>

      {tab === "history" ? <ClientHistory client={client} onOpenPeriod={(id) => { const i = client.periods.findIndex((p) => p.id === id); if (i >= 0) { setPIdx(i); setTab("period"); } }} /> : <>
      <PeriodBar client={client} pIdx={pIdx} period={period} onPick={selectPeriod} onIndex={setPIdx} numDays={numDays} onDays={onDays} />

      {!period && <div className="card" style={{ color: "var(--ink-3)" }}>Pick a month, week or day above to start uploading reports for <b>{client?.name}</b>.</div>}
      {period && (period._stub || loadingPeriod) && <div className="card" style={{ color: "var(--ink-3)" }}>Loading {period.label}…</div>}

      <div className="section-h"><h2>Scopes of work</h2><span className="rule" /></div>
      <div className="scopework">
        {activeScopes.map((sc) => { const col = scopeColor(sc.key); return (
          <div key={sc.key} className="scopework-card" style={{ borderLeftColor: col }}>
            <div style={{ display: "flex", gap: 11, alignItems: "center" }}>
              <span style={{ width: 36, height: 36, flex: "none", borderRadius: 11, background: col + "1c", color: col, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 10.5, letterSpacing: ".02em" }}>{scopeBadge(sc.key)}</span>
              <div style={{ minWidth: 0 }}>
                <div className="sw-name">{sc.name}</div>
                <div className="sw-sub">{sc.criteria ? `${sc.criteria.length} criteria` : `${sc.tasks.length} tasks`}{sc.tasks && sc.criteria ? ` · ${sc.tasks.length} tasks` : ""}{(SCOPE_REPORTS[sc.key] || []).length ? ` · ${SCOPE_REPORTS[sc.key].length} report${SCOPE_REPORTS[sc.key].length > 1 ? "s" : ""}` : ""}</div>
              </div>
            </div>
          </div>
        ); })}
        {activeScopes.length === 0 && <div style={{ color: "var(--ink-3)", fontSize: 13 }}>No scopes yet. Use “edit scopes” above.</div>}
      </div>

      {period && !period._stub && !loadingPeriod && <>
        <div className="row" style={{ justifyContent: "space-between", margin: "18px 0 6px", flexWrap: "wrap", gap: 8 }}>
          <div className="seg">
            <button className={"segbtn" + (viewFilter === "both" ? " on" : "")} onClick={() => setViewFilter("both")}>Both</button>
            <button className={"segbtn" + (viewFilter === "kpi" ? " on" : "")} onClick={() => setViewFilter("kpi")}>KPI</button>
            <button className={"segbtn" + (viewFilter === "sla" ? " on" : "")} onClick={() => setViewFilter("sla")}>SLA</button>
          </div>
        </div>

        {viewFilter !== "sla" && <>
          <div className="section-h"><h2>Reports for KPI</h2><span className="rule" /></div>
          {bulkSlots.length > 0 && can("data:entry") && <OneDriveBar me={me} client={client} period={period} onSaveOd={saveOd} picker={odPickRef} />}
          {bulkSlots.length > 1 && <BulkDrop count={bulkSlots.filter((k) => reps[k]).length} total={bulkSlots.length} busy={bulkBusy} onFiles={onBulkFiles} onCloud={cloud ? pickBulk : null} />}
          {activeScopes.filter((sc) => (SCOPE_REPORTS[sc.key] || []).length).map((sc) => (
            <div key={sc.key} style={{ marginBottom: 14 }}>
              <div style={{ fontSize: 12.5, fontWeight: 600, color: "var(--ink-2)", margin: "4px 0 8px" }}>{sc.name}</div>
              <div className="reportgrid">{scopeReportKeys(sc.key, client.scopes).map((rk) => <ReportCard key={rk} def={REPORT_BY_KEY[rk]} data={reps[rk]} onReport={onReport} onClear={clearReport} saved={!!client.colMap?.[rk]} onForget={forgetCols} onCloud={cloud ? pickOne : null} />)}</div>
            </div>
          ))}
          {activeScopes.every((sc) => !(SCOPE_REPORTS[sc.key] || []).length) && <div className="card" style={{ color: "var(--ink-3)" }}>The selected scopes don’t use uploaded reports for KPI. Enter their numbers below.</div>}
        </>}

        {viewFilter !== "kpi" && activeScopes.some((sc) => (SCOPE_SLA_REPORTS[sc.key] || []).length) && <>
          <div className="section-h"><h2>Reports for SLA</h2><span className="rule" /></div>
          <div style={{ fontSize: 11.5, color: "var(--ink-3)", margin: "-2px 0 8px" }}>Extra files needed only for SLA breakdowns. (The current-period AR inventory is taken from your KPI upload, so there is no need to re-upload it here.)</div>
          {activeScopes.filter((sc) => (SCOPE_SLA_REPORTS[sc.key] || []).length).map((sc) => (
            <div key={sc.key} style={{ marginBottom: 14 }}>
              <div style={{ fontSize: 12.5, fontWeight: 600, color: "var(--ink-2)", margin: "4px 0 8px" }}>{sc.name}</div>
              <div className="reportgrid">{SCOPE_SLA_REPORTS[sc.key].map((rk) => <ReportCard key={rk} def={REPORT_BY_KEY[rk]} data={reps[rk]} onReport={onReport} onClear={clearReport} saved={!!client.colMap?.[rk]} onForget={forgetCols} onCloud={cloud ? pickOne : null} />)}</div>
            </div>
          ))}
        </>}

        {viewFilter !== "sla" && activeScopes.some((s) => s.criteria) && <div className="section-h"><h2>KPI</h2><span className="rule" /></div>}
        {viewFilter !== "sla" && activeScopes.filter((sc) => sc.criteria).map((sc) => (
          <ScopeKpi key={sc.key} scope={sc} data={period.scopes?.[sc.key]} open={scopeOpen === sc.key} onToggle={() => setScopeOpen(scopeOpen === sc.key ? null : sc.key)} onSave={(data, method, file) => saveScopeData(sc.key, data, "saved " + sc.short, method, file)} />
        ))}
        {viewFilter !== "sla" && (reps.payment?.payerPay || reps.ar?.payerAR || reps.denial?.payerDenialCnt) && <>
          <div className="section-h"><h2>Payer analysis (Pareto)</h2><span className="rule" /></div>
          <ParetoSet reports={reps} scopes={activeScopes} />
        </>}
        {viewFilter !== "kpi" && (() => {
          const slaRows = activeScopes.filter((s) => s.criteria && s.sla?.length).flatMap((s) => s.sla);
          const ownSla = activeScopes.filter((s) => s.kind === "tasks");
          if (!slaRows.length && !ownSla.length) return null;
          const slaScope = { key: "sla", name: "Service Level Agreement", short: "SLA", kind: "tasks", tasks: slaRows };
          return <>
            <div className="section-h"><h2>Service Level Agreement (SLA)</h2><span className="rule" /></div>
            {slaRows.length > 0 && <ScopeTasks scope={slaScope} data={period.scopes?.sla} open={scopeOpen === "sla"} onToggle={() => setScopeOpen(scopeOpen === "sla" ? null : "sla")} onSave={(data) => saveScopeData("sla", data, "saved SLA")} />}
            {activeScopes.some((s) => s.key === "ar") && reps.ar?.aging && <TotalInventory data={reps.ar} />}
            {activeScopes.some((s) => s.key === "ar") && reps.arprev?.aging && <ARAgingComparison prev={reps.arprev} curr={reps.ar} />}
            {activeScopes.some((s) => s.key === "ar") && reps.ar?.statusBreakdown && <StatusBreakdown data={reps.ar} />}
            {activeScopes.some((s) => s.key === "ar") && reps.ar?.statusAging && <StatusAgingMatrix data={reps.ar} />}
            {activeScopes.some((s) => s.key === "ar") && reps.ar?.statusBreakdown && <CategoryBreakdown data={reps.ar} mapping={client.statusCategories} canEdit={can("scopes:edit")} onEdit={() => setCatEdit(true)} />}
            {ownSla.map((sc) => <ScopeTasks key={sc.key} scope={sc} data={period.scopes?.[sc.key]} open={scopeOpen === sc.key} onToggle={() => setScopeOpen(scopeOpen === sc.key ? null : sc.key)} onSave={(data) => saveScopeData(sc.key, data, "saved " + sc.short)} />)}
          </>;
        })()}
      </>}
      </>}
      {addScope && <AddScope current={client?.scopes || []} onClose={() => setAddScope(false)} onAdd={addScopeToClient} />}
      {editScopes && <EditScopes client={client} onClose={() => setEditScopes(false)} onSave={saveScopesList} />}
      {bulkPick && <BulkFolderPick folders={bulkPick.folders} onPick={pickBulkFolder} onCancel={() => setBulkPick(null)} />}
      {bulk && <BulkReview rows={bulk.rows} allowedKeys={bulkSlots} onCancel={() => setBulk(null)} onApply={applyBulk} title={bulk.meta?.title} desc={bulk.meta?.desc} onOtherFolder={bulk.meta?.source && cloud ? () => { setBulk(null); pickBulk(); } : null} />}
      {cloud && <OdPicker me={me} client={client} apiRef={odPickRef} onSaveOd={saveOd} />}
      {pending && <ColumnPicker pending={pending} onClose={nextInColQueue} onResolve={resolveColumns} />}
      {joinPending && <JoinKeyPicker req={joinPending} onClose={() => setJoinPending(null)} onApprove={resolveJoin} />}
      {catEdit && <CategoryMappingEditor statuses={Object.keys(repsRef.current.ar?.statusBreakdown || {}).sort()} mapping={client.statusCategories} onClose={() => setCatEdit(false)} onSave={saveCatMapping} />}
      <Confetti trigger={celebrate} />
    </div>
  );
}

/* ---- KPI scope: criteria rows ---- */
function ScopeKpi({ scope, data, open, onToggle, onSave }) {
  const init = () => { const o = {}; scope.criteria.forEach((c) => { o[c.key] = { ...(data?.crit?.[c.key] || {}) }; }); return o; };
  const [crit, setCrit] = useState(init);
  const [info, setInfo] = useState(null);
  useEffect(() => { setCrit(init()); }, [data]);
  const setCell = (k, patch) => setCrit((s) => ({ ...s, [k]: { ...s[k], ...patch } }));
  const filled = scope.criteria.filter((c) => { const v = crit[c.key]; return v && (v.value != null && v.value !== ""); }).length;
  const save = () => {
    const out = { crit: {} };
    for (const c of scope.criteria) { const cell = crit[c.key]; if (!cell) continue; if (["status", "yesno", "text"].includes(c.type)) { if (cell.value || cell.reason) out.crit[c.key] = { value: cell.value, method: cell.method || "manual", src: cell.src, inputs: cell.inputs, reason: cell.reason }; } else if (isNum(cell.value)) out.crit[c.key] = { ...cell }; else if (cell.method === "manual" && cell.calc && isNum(cell.calc.value)) out.crit[c.key] = { value: cell.calc.value, method: "upload", src: cell.calc.src, inputs: cell.calc.inputs }; }
    if (data?.aging) out.aging = data.aging;
    onSave(out, "manual", null);
  };
  return (
    <div className="scope-block">
      <div className="scope-head" onClick={onToggle}><div style={{ display: "flex", alignItems: "center", gap: 10 }}><span style={{ width: 10, height: 10, borderRadius: 3, background: scopeColor(scope.key), flex: "none", boxShadow: `0 1px 4px ${scopeColor(scope.key)}66` }} /><span className="name">{scope.name}</span></div><div className="row"><span className="meta">{filled}/{scope.criteria.length} filled</span><span style={{ color: "var(--ink-3)" }}>{open ? "▲" : "▼"}</span></div></div>
      {open && <div className="scope-body">
        <div className="krow khead"><span>Metric</span><span>Value</span><span>Source</span><span style={{ textAlign: "right" }}>Status</span></div>
        {scope.criteria.map((c) => <CriterionRow key={c.key} crit={c} cell={crit[c.key]} onChange={(patch) => setCell(c.key, patch)} onInfo={() => setInfo({ crit: c, cell: crit[c.key], rule: c.rule })} />)}
        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 14 }}><button className="btn primary" onClick={save}>Save {scope.short}</button></div>
      </div>}
      {info && <InfoModal data={info} onClose={() => setInfo(null)} />}
    </div>
  );
}

function CriterionRow({ crit, cell, onChange, onInfo }) {
  const rule = crit.rule;
  const val = cell?.value;
  const pctToInput = (v) => (v == null || v === "" ? "" : +(v * 100).toFixed(2));
  const isPct = crit.type === "pct" || (crit.type === "calc" && crit.fmt === FMT.PERCENT);
  const isDays = crit.type === "calc" && crit.fmt === FMT.DAYS;
  const status = ["calc", "pct", "count"].includes(crit.type) ? statusFor(val, rule) : critStatus(crit, cell, rule);
  const showChip = (isNum(val) && ["calc", "pct", "count"].includes(crit.type)) || (["status", "yesno"].includes(crit.type) && val);
  const src = cell?.method === "upload" ? "auto" : cell?.method === "manual" ? "manual" : (val != null && val !== "") ? "" : "enter";
  const num = (raw) => (n(raw) == null ? undefined : n(raw));
  const notMet = showChip && (status === "warn" || status === "miss");
  return (
    <>
    <div className="krow">
      <div><div className="cn">{crit.name}{onInfo && <button className="infobtn" title="How this is calculated" onClick={onInfo}>ⓘ</button>}</div>{crit.std && <div className="cstd">{crit.std}</div>}</div>
      <div className="cin">
        {isPct && <span className="unitbox"><input className="mono" value={pctToInput(val)} placeholder="0" onChange={(e) => onChange({ value: n(e.target.value) == null ? undefined : n(e.target.value) / 100, method: "manual" })} /><i>%</i></span>}
        {isDays && <span className="unitbox"><input className="mono" value={val == null || val === "" ? "" : Math.round(val)} placeholder="0" onChange={(e) => onChange({ value: num(e.target.value), method: "manual" })} /><i>days</i></span>}
        {crit.type === "count" && <input className="mono" value={val ?? ""} placeholder="0" onChange={(e) => onChange({ value: num(e.target.value), method: "manual" })} />}
        {crit.type === "calc" && !isPct && !isDays && <input className="mono" value={val ?? ""} placeholder="0" onChange={(e) => onChange({ value: num(e.target.value), method: "manual" })} />}
        {crit.type === "status" && <select value={val || ""} onChange={(e) => onChange({ value: e.target.value || undefined, method: "manual" })}><option value="">—</option><option>Met</option><option>Not Met</option><option>N/A</option></select>}
        {crit.type === "yesno" && <select value={val || ""} onChange={(e) => onChange({ value: e.target.value || undefined, method: "manual" })}><option value="">—</option><option>Yes</option><option>No</option></select>}
        {crit.type === "text" && <input value={val || ""} onChange={(e) => onChange({ value: e.target.value || undefined, method: "manual" })} placeholder="—" />}
      </div>
      <div className="csrc">{src === "auto" ? <span className="tag up">auto</span> : src === "manual" ? <span className="tag man">manual</span> : src === "enter" ? <span style={{ color: "var(--ink-3)", fontSize: 11 }}>enter</span> : null}</div>
      <div style={{ textAlign: "right" }}>{showChip && status !== "info" && <span className={"chip " + status}>{stLabel(status)}</span>}</div>
    </div>
    {cell?.method === "manual" && cell?.calc && cell.calc.value != null && cell.calc.value !== val && <div className="calcrow"><span>Your typed value is in use. Calculated from the reports: <b>{typeof cell.calc.value === "number" ? fmtVal(cell.calc.value, crit.fmt || rule?.fmt) : String(cell.calc.value)}</b></span><button className="link" onClick={() => onChange({ value: cell.calc.value, method: "upload", src: cell.calc.src, inputs: cell.calc.inputs, calc: undefined })}>Use the calculated value</button></div>}
    {notMet && <div className="reasonrow"><input value={cell?.reason || ""} onChange={(e) => onChange({ reason: e.target.value || undefined })} placeholder="Not met, add reason (optional)" /></div>}
    </>
  );
}
const AGING_BUCKETS = ["A: 0-30 Days", "B: 31-60 Days", "C: 61-90 Days", "D: 91-120 Days", "E: 121-180 Days", "F: 181-360 Days", "G: 360 Plus"];
function AgingEditor({ aging, setAging }) {
  const [on, setOn] = useState(aging.rows.some((r) => r.start || r.end));
  const setRow = (i, k, v) => { const rows = aging.rows.map((r, idx) => idx === i ? { ...r, [k]: v } : r); setAging({ ...aging, rows }); };
  return (
    <div style={{ marginTop: 14, borderTop: "1px solid var(--line-2)", paddingTop: 12 }}>
      <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 500, cursor: "pointer" }}><input type="checkbox" checked={on} onChange={(e) => setOn(e.target.checked)} /> Record AR aging comparison (optional)</label>
      {on && <div style={{ marginTop: 10 }}>
        <div className="row" style={{ marginBottom: 8 }}><span style={{ fontSize: 12, color: "var(--ink-2)" }}>Compare by</span><select value={aging.scope} onChange={(e) => setAging({ ...aging, scope: e.target.value })}><option value="month">Month (start vs end)</option><option value="week">Week (start vs end)</option></select></div>
        <div className="panel"><table><thead><tr><th>Aging bucket</th><th>Start of {aging.scope}</th><th>End of {aging.scope}</th></tr></thead>
          <tbody>{aging.rows.map((r, i) => (<tr key={i}><td style={{ color: "var(--ink-2)" }}>{r.bucket}</td><td><input className="mono" value={r.start} onChange={(e) => setRow(i, "start", e.target.value)} placeholder="$" /></td><td><input className="mono" value={r.end} onChange={(e) => setRow(i, "end", e.target.value)} placeholder="$" /></td></tr>))}</tbody></table></div>
      </div>}
    </div>
  );
}

/* ---- tasks scope: SLA / Credentialing / PCC ---- */
function ScopeTasks({ scope, data, open, onToggle, onSave }) {
  const [rows, setRows] = useState(() => scope.tasks.map((t) => { const ex = (data?.tasks || []).find((r) => r.name === t); return ex || { name: t, target: "", received: "", achieved: "", comment: "" }; }));
  const savedCount = (data?.tasks || []).filter((r) => r.target || r.received || r.achieved || r.comment).length;
  const setRow = (i, k, v) => setRows((s) => s.map((r, idx) => idx === i ? { ...r, [k]: v } : r));
  const save = () => { onSave({ tasks: rows }); onToggle(); };
  return (
    <div className="scope-block">
      <div className="scope-head" onClick={onToggle}><div style={{ display: "flex", alignItems: "center", gap: 10 }}><span style={{ width: 10, height: 10, borderRadius: 3, background: scopeColor(scope.key), flex: "none", boxShadow: `0 1px 4px ${scopeColor(scope.key)}66` }} /><span className="name">{scope.name}</span></div><div className="row"><span className="meta">{savedCount}/{scope.tasks.length} filled</span><span style={{ color: "var(--ink-3)" }}>{open ? "▲" : "▼"}</span></div></div>
      {open && <div className="scope-body">
        <div className="panel" style={{ marginTop: 8 }}><table className="tasktbl"><thead><tr><th style={{ width: "26%" }}>Metric</th><th>Target</th><th>Received</th><th>Achieved</th><th style={{ width: "28%" }}>Comment</th></tr></thead>
          <tbody>{rows.map((r, i) => (<tr key={i}><td style={{ color: "var(--ink)", fontWeight: 500 }}>{r.name}</td>
            <td><input value={r.target} onChange={(e) => setRow(i, "target", e.target.value)} placeholder="e.g. 725 / Based on project" /></td>
            <td><input className="mono" value={r.received} onChange={(e) => setRow(i, "received", e.target.value)} placeholder="0" /></td>
            <td><input className="mono" value={r.achieved} onChange={(e) => setRow(i, "achieved", e.target.value)} placeholder="0" /></td>
            <td><input value={r.comment} onChange={(e) => setRow(i, "comment", e.target.value)} placeholder="—" /></td></tr>))}</tbody></table></div>
        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 14 }}><button className="btn primary" onClick={save}>Save {scope.short} to period</button></div>
      </div>}
    </div>
  );
}

/* ---- four source-report upload cards ---- */
function ReportsPanel({ reps, onReport, onAvg3, onClear }) {
  return (<div className="reportgrid">{REPORT_DEFS.map((d) => <ReportCard key={d.key} def={d} data={reps[d.key]} onReport={onReport} onAvg3={onAvg3} onClear={onClear} />)}</div>);
}
function ReportCard({ def, data, onReport, onAvg3, onClear, saved, onForget, onCloud }) {
  const [busy, setBusy] = useState(false);
  const [drag, setDrag] = useState(false);
  const fileRef = useRef(null);
  const [err, setErr] = useState("");
  const cloud = !!onCloud; // OneDrive switched on: the file is picked there, never from the computer
  const readFile = async (f, src) => { if (!f) return; setBusy(true); setErr(""); try { const rep = await readReport(f); onReport(def.key, rep, src); } catch { setErr("That file could not be read. Try saving it again as .xlsx or .csv."); } setBusy(false); };
  const upload = async (e) => { const f = e.target.files?.[0]; e.target.value = ""; await readFile(f); };
  const onDrop = async (e) => { e.preventDefault(); setDrag(false); const f = e.dataTransfer?.files?.[0]; await readFile(f); };
  const dragProps = cloud ? {} : { onDragOver: (e) => { e.preventDefault(); setDrag(true); }, onDragLeave: () => setDrag(false), onDrop };
  const choose = () => { setErr(""); if (cloud) onCloud(def, readFile); else if (fileRef.current) fileRef.current.click(); };
  const dollar = (v) => (v != null ? "$" + Math.round(v).toLocaleString() : null);
  const figs = [];
  if (data) {
    if (def.key === "payment") { figs.push(["Payments", dollar(data.payments)]); figs.push(["Charges", dollar(data.charges)]); figs.push(["Contractual adj.", dollar(data.adjustments)]); if (data.patientPaid != null) figs.push(["Patient / Insurance", `${dollar(data.patientPaid)} / ${dollar(data.insurancePaid)}`]); }
    if (def.key === "ar") { figs.push(["Total AR", dollar(data.totalAR)]); figs.push(["Charges", dollar(data.totalCharges)]); if (data.aging) figs.push(["Aging", `${Object.keys(data.aging).length} buckets · ${data.agingSource}`]); if (data.appeals != null) figs.push(["Appeals found", String(data.appeals)]); }
    if (def.key === "denial") figs.push(["Denial count", String(data.count)]);
    if (def.key === "submission") figs.push(["Submissions", String(data.count)]);
    if (def.key === "audit") figs.push(["Avg audit score", isNum(data.score) ? data.score.toFixed(1) + "%" : "not detected"]);
    if (def.key === "escalation") figs.push(["Escalations", String(data.count)]);
    if (def.key === "ev") { if (data.evTotal != null) figs.push(["Completed / Total", `${data.evCompleted} / ${data.evTotal}`]); if (data.evUrgent != null) figs.push(["Urgent / Regular", `${data.evUrgent} / ${data.evRegular}`]); }
    if (def.key === "paysub") figs.push(["Rows · date col", `${data.count} · ${data.dateIdx != null ? "found" : "not detected"}`]);
    if (def.key === "paypaid") figs.push(["Rows · date col", `${data.count} · ${data.dateIdx != null ? "found" : "not detected"}`]);
    if (def.key === "ledger") { if (data.ledgerAllowed != null) { figs.push(["Total allowed", dollar(data.ledgerAllowed)]); figs.push(["Underpaid $", dollar(data.ledgerUnder)]); } }
    if (def.key === "postrecv") { if (data.postingIn != null) figs.push(["Received for posting $", dollar(data.postingIn)]); }
    if (def.key === "variance") { if (data.varianceAmt != null) figs.push(["Discrepancy $", dollar(data.varianceAmt)]); }
    if (def.key === "noncontr") { if (data.writeoff != null) figs.push(["Write-off $ (excl CO45)", dollar(data.writeoff)]); }
    if (def.key === "recoup") { if (data.recoupAmt != null) figs.push(["Recouped $", dollar(data.recoupAmt)]); }
    if (def.key === "pcccalls") { if (data.callsTotal != null) { figs.push(["Total inflow", String(data.callsTotal)]); figs.push(["Picked / Missed / Dropped", `${data.callsPicked} / ${data.callsMissed} / ${data.callsDropped}`]); if (data.callsOther) figs.push(["Unclassified", String(data.callsOther)]); } }
    if (def.key === "pcccallback") { if (data.cbTotal != null) figs.push(["Callbacks made / logged", `${data.cbDone} / ${data.cbTotal}`]); }
    if (def.key === "pcccorr") { if (data.corrTotal != null) figs.push(["Sent / Received", `${data.corrDone} / ${data.corrTotal}`]); }
    if (def.key === "pccfax") { if (data.faxTotal != null) figs.push(["Completed / Received", `${data.faxDone} / ${data.faxTotal}`]); }
  }
  return (
    <div className={"report-card" + (drag ? " dragover" : "")} {...dragProps}>
      <div className="rc-top"><div><div className="rc-name">{def.name}</div><div className="rc-hint">{def.hint}</div></div>{data && <span className="tag up">loaded</span>}</div>
      {err && <div className="rc-err">{err}</div>}
      {!cloud && <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={upload} />}
      {!data ? <div style={{ marginTop: 12 }}><button className="btn sm" onClick={choose} disabled={busy}>{busy ? "Reading…" : cloud ? "Upload from OneDrive" : "Upload file"}</button>{!cloud && <div style={{ fontSize: 10.5, color: "var(--ink-3)", marginTop: 6 }}>{drag ? "Drop to upload" : "or drag and drop a file here"}</div>}</div> : <>
        <div className="rc-figs">{figs.map(([k, v], i) => <div key={i} className="rc-fig"><span className="k">{k}</span>{v == null ? <span className="v miss">not detected</span> : <span className="v mono">{v}</span>}</div>)}
          <div className="rc-fig"><span className="k">Rows read</span><span className="v mono">{data.rows}</span></div>
        </div>
        {data.cols && data.cols.length > 0 && <div className="rc-cols"><div className="rc-cols-h">Columns read</div>{data.cols.map((c) => <div key={c.id} className="rc-col"><span>{c.label}</span><b title={c.name}>{c.name}</b></div>)}</div>}
        {saved && <div className="rc-saved">Using the columns you chose earlier. <button className="link" onClick={() => onForget(def.key)}>Forget them</button></div>}
        <div className="rc-file">{data.fileName}</div>
        <div className="row" style={{ marginTop: 10 }}><button className="btn sm" onClick={choose} disabled={busy}>{busy ? "Reading…" : cloud ? "Replace from OneDrive" : "Replace"}</button><button className="btn sm ghost" onClick={() => onClear(def.key)}>Clear</button></div>
      </>}
    </div>
  );
}

/* ============================== clients =============================== */
function Clients({ me, can, clients, users, refreshClients }) {
  const [addOpen, setAddOpen] = useState(false);
  const [openId, setOpenId] = useState(null);
  const [delFor, setDelFor] = useState(null);
  const [handFor, setHandFor] = useState(null);
  const [notice, setNotice] = useState("");
  const [showArch, setShowArch] = useState(false);
  const archived = REMOTE ? clients.filter((c) => c.archived) : [];
  const restore = async (c) => { try { await api("client.restore", { id: c.id }); await refreshClients(); setNotice(""); } catch (e) { setNotice(e.message); } };
  const showOwner = REMOTE && (me.role === "admin" || me.role === "manager");
  const handOver = async (c, ownerId) => {
    try { await api("client.reassign", { id: c.id, ownerId }); await refreshClients(); setHandFor(null); setNotice(""); }
    catch (e) { setNotice(e.message); setHandFor(null); }
  };
  const sub = !REMOTE ? "Open a client to upload reports and see its KPIs & SLA"
    : me.role === "admin" ? "Every client, and who maintains it"
    : me.role === "manager" ? "Your clients and your team's clients"
    : "The clients you maintain. Only you, your manager and the Admin can see them.";
  const add = async ({ name, scopes }) => {
    const c = { id: uid(), name: name.trim(), scopes, rules: {}, periods: [], createdAt: nowIso(), createdBy: me.username };
    await sSet(clientKey(c.id), c); await sSet(K.clients, [...clients, { id: c.id, name: c.name, scopes }]);
    await appendAudit({ user: me.username, role: me.role, action: `added client (${scopes.map((s) => SCOPE_BY_KEY[s].short).join(", ")})`, client: c.name });
    await refreshClients(); setAddOpen(false); setOpenId(c.id);
  };
  const del = async (id, name) => {
    try { await store().delete(clientKey(id)); } catch {}
    await sSet(K.clients, clients.filter((c) => c.id !== id));
    if (!REMOTE) await appendAudit({ user: me.username, role: me.role, action: "removed client", client: name });
    await refreshClients(); setDelFor(null);
  };

  if (openId) return <ClientWorkspace me={me} can={can} clientId={openId} onBack={() => setOpenId(null)} onScopesChanged={refreshClients} />;
  const live = clients.filter((c) => !c.archived);
  return (
    <div className="wrap">
      <div className="pagehead"><div><h1>Clients</h1><div className="sub">{sub}</div></div>{can("clients:add") && <button className="btn primary" onClick={() => setAddOpen(true)}>+ Add client</button>}</div>
      {notice && <div className="rc-err" style={{ marginBottom: 10 }}>{notice}</div>}
      {live.length === 0 ? <div className="empty"><p>{archived.length ? "No active clients. Removed clients are listed below." : "No clients yet. Add one to get started. You’ll pick its scopes of work, then upload reports inside it."}</p></div> :
        <div className="panel"><table><thead><tr><th>Client</th><th>Scopes of work</th>{showOwner && <th>Maintained by</th>}<th></th></tr></thead>
          <tbody>{live.map((c) => (<tr key={c.id} style={{ cursor: "pointer" }} onClick={() => setOpenId(c.id)}>
            <td style={{ color: "var(--ink)", fontWeight: 600 }}><div style={{ display: "flex", alignItems: "center", gap: 11 }}><span style={{ width: 34, height: 34, flex: "none", borderRadius: 11, background: "linear-gradient(135deg,var(--brand1),var(--brand3))", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 13, boxShadow: "0 3px 8px -2px rgba(44,201,176,.5)" }}>{c.name.trim().charAt(0).toUpperCase()}</span>{c.name}</div></td>
            <td><div className="chips">{(c.scopes || []).map((s) => <span key={s} className="pill" style={{ background: scopeColor(s) + "1c", color: scopeColor(s) }}>{SCOPE_BY_KEY[s]?.short}</span>)}</div></td>
            {showOwner && <td><span className="ownertag">{c.ownerId === me.id ? "You" : c.ownerName || "Unknown"}</span></td>}
            <td style={{ textAlign: "right", whiteSpace: "nowrap" }} onClick={(e) => e.stopPropagation()}><button className="btn sm" onClick={() => setOpenId(c.id)}>Open</button>{" "}{REMOTE && can("clients:reassign") && <><button className="btn sm ghost" onClick={() => setHandFor(c)}>Hand over</button>{" "}</>}{can("clients:delete") && <button className="btn sm danger" title={"Remove " + c.name} onClick={() => setDelFor(c)}>Remove</button>}</td>
          </tr>))}</tbody>
        </table></div>}
      {addOpen && <AddClient onClose={() => setAddOpen(false)} onCreate={add} />}
      {archived.length > 0 && <div className="arch">
        <button className="link" onClick={() => setShowArch(!showArch)}>{showArch ? "Hide" : "Show"} removed clients ({archived.length})</button>
        {showArch && <div className="panel" style={{ marginTop: 10 }}><table><thead><tr><th>Removed client</th><th>Maintained by</th><th>Removed</th><th></th></tr></thead>
          <tbody>{archived.map((c) => <tr key={c.id}><td style={{ color: "var(--ink)", fontWeight: 600 }}>{c.name}</td><td>{c.ownerName || "Unknown"}</td><td style={{ fontSize: 12 }}>{c.archivedAt ? fmtDateTime(c.archivedAt) : ""}{c.archivedBy ? " by " + c.archivedBy : ""}</td><td style={{ textAlign: "right" }}><button className="btn sm" onClick={() => restore(c)}>Restore</button></td></tr>)}</tbody>
        </table><div className="arch-note">Removed clients keep all their weeks and months. Restoring brings everything back for the person who maintains it.</div></div>}
      </div>}
      {delFor && <RemoveClient client={delFor} onClose={() => setDelFor(null)} onConfirm={() => del(delFor.id, delFor.name)} />}
      {handFor && <HandOver client={handFor} me={me} users={users || []} onClose={() => setHandFor(null)} onSave={(ownerId) => handOver(handFor, ownerId)} />}
    </div>
  );
}

// a Manager hands a client to anyone in their team (or keeps it); the Admin can hand it to anyone
function HandOver({ client, me, users, onClose, onSave }) {
  const options = users.filter((u) => u.id !== client.ownerId && (me.role === "admin" || u.id === me.id || u.managerId));
  const [to, setTo] = useState(""); const [busy, setBusy] = useState(false);
  const label = (u) => (u.id === me.id ? "Me (" + (u.name || u.username) + ")" : (u.name || u.username)) + " · " + roleLabel(u.role);
  return (<Modal onClose={onClose} title={"Hand over " + client.name} desc="Choose who maintains this client from now on. The person you pick sees it straight away, with all its periods and numbers.">
    <div className="field"><label>Currently maintained by</label><div style={{ fontSize: 13, fontWeight: 600 }}>{client.ownerId === me.id ? "You" : client.ownerName || "Unknown"}</div></div>
    <div className="field"><label>Hand it to</label>
      {options.length ? <select autoFocus value={to} onChange={(e) => setTo(e.target.value)}><option value="">Choose…</option>{options.map((u) => <option key={u.id} value={u.id}>{label(u)}</option>)}</select>
        : <div style={{ fontSize: 12.5, color: "var(--ink-3)" }}>There is nobody else to hand it to yet. {me.role === "manager" ? "People appear here once they create an account with you as their reporting manager." : ""}</div>}
    </div>
    <Actions onClose={onClose} onOk={async () => { setBusy(true); await onSave(to); }} okLabel={busy ? "Handing over…" : "Hand over"} okDisabled={!to || busy} />
  </Modal>);
}

function RemoveClient({ client, onClose, onConfirm }) {
  const [typed, setTyped] = useState("");
  const [busy, setBusy] = useState(false);
  const [periods, setPeriods] = useState(null);
  useEffect(() => { (async () => { const full = await sGet(clientKey(client.id)); setPeriods(full?.periods?.length ?? 0); })(); }, [client.id]);
  const match = typed.trim().toLowerCase() === client.name.trim().toLowerCase();
  const go = async () => { if (!match || busy) return; setBusy(true); await onConfirm(); };
  return (
    <Modal title={"Remove " + client.name + "?"} desc={REMOTE ? "It disappears from everyone's lists. Its history is kept." : "This cannot be undone."} onClose={onClose}>
      {REMOTE ? <div className="note-box">
        {client.name} disappears from Clients and the Overview for everyone. All {periods == null ? "its" : periods === 1 ? "of its 1 saved period" : `${periods} of its saved periods`} stay in the Google Sheet, and the Admin can restore it at any time from the Clients tab.
      </div> : <div className="note-box" style={{ background: "var(--miss-bg)", color: "#8a2b2b" }}>
        Removing this client permanently deletes {periods == null ? "its" : periods === 0 ? "its" : periods === 1 ? "its 1 recorded period" : `all ${periods} recorded periods`} along with every uploaded report, KPI value, SLA entry and reason note stored for it. Its scores will disappear from the Overview.
      </div>}
      <div className="field">
        <label>Type <b>{client.name}</b> to confirm</label>
        <input autoFocus value={typed} onChange={(e) => setTyped(e.target.value)} placeholder={client.name} onKeyDown={(e) => { if (e.key === "Enter") go(); }} />
      </div>
      <div className="modal-actions">
        <button className="btn ghost" onClick={onClose}>Cancel</button>
        <button className="btn danger solid" disabled={!match || busy} onClick={go}>{busy ? "Removing…" : "Remove client"}</button>
      </div>
    </Modal>
  );
}

/* ===================== formulas and definitions ===================== */
function Formulas() {
  const [q, setQ] = useState("");
  const [pick, setPick] = useState("all");
  const needle = q.trim().toLowerCase();
  const scopes = SCOPES.filter((s) => pick === "all" || s.key === pick);
  const rowsFor = (sc) => (sc.criteria || []).map((c) => {
    const b = FORMULA_BOOK[sc.key + "." + c.key] || {};
    return { c, b, manual: !!b.manual, formula: b.page || c.formula || null, from: b.from || TYPED, bands: ruleBands(c.rule) };
  }).filter((r) => !needle || (r.c.name + " " + (r.formula || "") + " " + r.from).toLowerCase().includes(needle));
  const slaFor = (sc) => [...(sc.sla || []), ...(sc.tasks || [])].filter((t) => !needle || t.toLowerCase().includes(needle));
  const shown = scopes.map((sc) => ({ sc, rows: rowsFor(sc), sla: slaFor(sc) })).filter((x) => x.rows.length || x.sla.length);
  return (
    <div className="wrap">
      <div className="pagehead"><div><h1>Formulas and definitions</h1><div className="sub">How every KPI and SLA number is worked out, by scope of work</div></div></div>
      <div className="fm-tools">
        <input className="fm-search" value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search a KPI, for example Denial Rate" />
        <div className="fm-chips">
          <button className={"fm-chip" + (pick === "all" ? " on" : "")} onClick={() => setPick("all")}>All scopes</button>
          {SCOPES.map((s) => <button key={s.key} className={"fm-chip" + (pick === s.key ? " on" : "")} onClick={() => setPick(s.key)} style={pick === s.key ? { background: scopeColor(s.key), borderColor: scopeColor(s.key) } : undefined}>{s.name}</button>)}
        </div>
      </div>

      {pick === "all" && !needle && <div className="panel fm-how">
        <div className="fm-how-h">How the scores are worked out</div>
        <div className="fm-how-grid">
          <div><b>Each KPI</b><p>Compared with its target and marked green (on target), amber (at risk) or red (missed).</p></div>
          <div><b>Each SLA row</b><p>Achieved ÷ Received × 100. Green at 90% or above, amber from 70% to just under 90%, red below 70%.</p></div>
          <div><b>A client's overall score</b><p>Metrics on target ÷ metrics that have a value × 100. Anything left blank is not counted.</p></div>
          <div><b>Overall health</b><p>Healthy at 90% or above, at risk from 80% to just under 90%, below target under 80%. Clients at 90% or above get the shout-out.</p></div>
        </div>
        <div className="fm-note">Targets shown here are the standard ones. A client can be given its own target for any KPI from that client's page, and that client is then scored against its own target.</div>
      </div>}

      {shown.length === 0 && <div className="card" style={{ color: "var(--ink-3)" }}>Nothing matches that search.</div>}
      {shown.map(({ sc, rows, sla }) => (
        <div key={sc.key} className="panel fm-scope">
          <div className="fm-scope-h">
            <span className="fm-badge" style={{ background: scopeColor(sc.key) + "26", color: scopeColor(sc.key) }}>{scopeBadge(sc.key)}</span>
            <div><div className="fm-scope-name">{sc.name}</div><div className="fm-scope-sub">{rows.length} KPI{rows.length === 1 ? "" : "s"}{sla.length ? ` · ${sla.length} SLA row${sla.length === 1 ? "" : "s"}` : ""}</div></div>
          </div>
          {rows.length > 0 && <div className="fm-table">
            <div className="fm-row fm-head"><div>KPI</div><div>Formula</div><div>Target</div><div>Numbers come from</div></div>
            {rows.map(({ c, b, manual, formula, from, bands }) => (
              <div key={c.key} className="fm-row">
                <div className="fm-kpi"><span>{c.name}</span><span className={"fm-tag " + (manual ? "man" : "auto")}>{manual ? "Typed in" : "From upload"}</span></div>
                <div>{formula ? <div className="fm-formula">{formula}</div> : <div className="fm-plain">No calculation. The value is entered as it is.</div>}{b.alt && <div className="fm-alt">{b.alt}</div>}</div>
                <div className="fm-bands">{bands.length ? bands.map((x, i) => <span key={i} className={"chip " + x.cls}>{x.t}</span>)
                  : c.type === "status" ? <><span className="chip ok">Met{c.std ? ": within " + c.std : ""}</span><span className="chip miss">Not met</span></>
                  : c.type === "yesno" ? <><span className="chip ok">Yes</span><span className="chip miss">No</span></>
                  : <span className="fm-plain">Tracked only, not scored</span>}</div>
                <div className="fm-from">{from}</div>
              </div>
            ))}
          </div>}
          {sla.length > 0 && <div className="fm-sla">
            <div className="fm-sla-h">SLA rows</div>
            <div className="fm-sla-list">{sla.map((t) => <span key={t} className="fm-sla-item">{t}</span>)}</div>
            <div className="fm-sla-f"><span className="fm-formula">Achieved ÷ Received × 100</span><span className="chip ok">90% or above</span><span className="chip warn">70% to just under 90%</span><span className="chip miss">Below 70%</span><span className="fm-from">Typed in by the supervisor</span></div>
          </div>}
        </div>
      ))}
    </div>
  );
}

/* ============================ user guide ============================ */
const usingBrowserStore = () => store() === localStore;
const GUIDE_TOC = [
  ["start", "What Drishti does"], ["signin", "Signing in"], ["screens", "The screens"], ["client", "Step 1: Add a client"],
  ["duration", "Step 2: Pick the period"], ["upload", "Step 3: Upload the reports"], ["onedrive", "Uploading from OneDrive"], ["files", "Which files, and what they need"],
  ["columns", "How Drishti finds the right column"], ["popups", "When Drishti asks you something"], ["typed", "Step 4: Numbers you type in"], ["reading", "Step 5: Read the numbers"],
  ["overview", "The Overview"], ["history", "History: past months and weeks"], ["targets", "Targets"], ...(REMOTE ? [["access", "Who sees which clients"]] : []), ["manage", REMOTE ? "People, periods and the audit log" : "Users, periods and the audit log"], ["feedback", "Questions and feedback"],
  ["routine", "Your routine each period"], ["help", "If something looks wrong"],
];
function GuideSec({ id, n, title, lead, children }) {
  return (<section id={"gd-" + id} className="panel gd-sec"><div className="gd-h">{n != null && <span className="gd-num">{n}</span>}<div><h2>{title}</h2>{lead && <p className="gd-lead">{lead}</p>}</div></div>{children}</section>);
}
const GTip = ({ children }) => <div className="gd-note"><b>Good to know</b>{children}</div>;
const GWarn = ({ children }) => <div className="gd-note warn"><b>Take care</b>{children}</div>;
function Guide({ me, can, onGo }) {
  const local = usingBrowserStore();
  const cloud = msSupported(); // OneDrive switched on: every upload is picked in OneDrive
  const odAcct = odAccountShown();
  const go = (id) => { const el = document.getElementById("gd-" + id); if (el) el.scrollIntoView({ behavior: "smooth", block: "start" }); };
  const scopesWithReports = SCOPES.filter((s) => (SCOPE_REPORTS[s.key] || []).length || (SCOPE_SLA_REPORTS[s.key] || []).length);
  return (
    <div className="wrap">
      <div className="pagehead"><div><h1>User guide</h1><div className="sub">Everything you need to use Drishti, from signing in to reading the scores. No earlier knowledge needed.</div></div></div>
      <div className="gd-layout">
        <nav className="gd-toc panel">
          <div className="gd-toc-h">On this page</div>
          {GUIDE_TOC.map(([id, t]) => <button key={id} onClick={() => go(id)}>{t}</button>)}
        </nav>
        <div className="gd-body">

          <GuideSec id="start" title="What Drishti does" lead="Drishti shows, for every client, whether we are meeting the KPI and SLA targets we agreed, and names what is holding a client back.">
            <p>You give it the client's reports for a period. It reads them, works out each KPI with the agreed formula, colours it green, amber or red against the target, and rolls everything up into one score per client. Leadership sees all clients side by side on the Overview.</p>
            <div className="gd-flow"><span>Add the client</span><i>→</i><span>Pick the month or week</span><i>→</i><span>Upload the reports</span><i>→</i><span>Type in the rest</span><i>→</i><span>Read the scores</span></div>
            <GTip>The exact formula and target for every KPI is on the <button className="link" onClick={() => onGo("formulas")}>Formulas</button> tab.</GTip>
          </GuideSec>

          {REMOTE ? <GuideSec id="signin" title="Signing in" lead="The first screen has two doors: Sign in as Admin and Sign in as User.">
            <ol className="gd-steps">
              <li><b>Sign in as Admin</b> is for the Admin only. That account is already set up, and nobody can create another Admin.</li>
              <li><b>Everyone else uses Sign in as User.</b> The first time, click <em>Create account</em> and fill in your full name, your Purview email, your role (Supervisor, Team Lead, SME or Manager), your <b>reporting manager</b>, and a password of at least 6 characters, twice. The reporting manager is required.</li>
              <li><b>Managers</b> choose <em>Manager</em> as their role. Your email must already be on the managers list the Admin keeps, and your reporting manager is the Admin. If Drishti says you are not on the list, ask the Admin to add you first.</li>
              <li><b>Then sign in</b> with your email and password. Drishti keeps you signed in on that computer for 12 hours. To leave, click <em>Sign out</em> at the top right.</li>
            </ol>
            <GTip>Forgot your password? Only the Admin can set a new one, from the People tab with <em>Reset password</em>.</GTip>
            <h3>What each role can do</h3>
            <div className="gd-table">
              <div className="gd-tr gd-th"><span>Role</span><span>Can do</span></div>
              <div className="gd-tr"><span><b>Supervisor, Team Lead, SME</b></span><span>Work on the clients they maintain: add clients, upload reports, type in values, change a client's scopes. They do not see anyone else's clients.</span></div>
              <div className="gd-tr"><span><b>Manager</b></span><span>Everything above for their own clients and for every client maintained by the people who report to them. Hand a client from one team member to another, change targets, and see the My team tab and their team's audit log.</span></div>
              <div className="gd-tr"><span><b>Admin</b></span><span>Everything, for every client. Add and remove managers, change anyone's reporting manager, reset passwords, remove accounts and remove a period.</span></div>
            </div>
            <p className="gd-you">You are signed in as <b>{me.name || me.username}</b>, role <b>{roleLabel(me.role)}</b>{me.managerName ? <>, reporting to <b>{me.managerName}</b></> : null}.</p>
          </GuideSec> : <GuideSec id="signin" title="Signing in" lead="You create your account once. After that you only sign in.">
            <ol className="gd-steps">
              <li><b>Create your account.</b> On the first screen enter your Purview email, choose a password of at least 6 characters, type it again to confirm, and click <em>Create account</em>.</li>
              <li><b>Sign in.</b> Your email is already filled in. Enter your password and click <em>Sign in</em>.</li>
              <li><b>Next time</b> you go straight to the sign in screen. To leave, click <em>Sign out</em> at the top right.</li>
            </ol>
            <GTip>Forgot your password? Ask an Admin. They can set a new one for you from the Users tab with <em>Reset password</em>.</GTip>
            {local && <GWarn>What you enter is saved in this browser on this computer. Use the same computer and the same browser each time. A private or incognito window, or clearing the browser's site data, will show an empty Drishti.</GWarn>}
            <h3>What each role can do</h3>
            <div className="gd-table">
              <div className="gd-tr gd-th"><span>Role</span><span>Can do</span></div>
              <div className="gd-tr"><span><b>Supervisor</b></span><span>Add and remove clients, change a client's scopes, upload reports, type in values, use Formulas, Guide, Q&amp;A and Feedback.</span></div>
              <div className="gd-tr"><span><b>Manager</b></span><span>Everything a Supervisor can, plus change a client's targets and see the Audit log.</span></div>
              <div className="gd-tr"><span><b>Admin</b></span><span>Everything a Manager can, plus add, remove and reset users, and remove a period.</span></div>
            </div>
            <p>The Admin role is by name only. Everyone who creates an account is a Supervisor. An Admin can make someone a Manager from the Users tab.</p>
            <p className="gd-you">You are signed in as <b>{me.username}</b>, role <b>{roleLabel(me.role)}</b>.</p>
          </GuideSec>}

          <GuideSec id="screens" title="The screens" lead="The tabs along the top take you everywhere.">
            <div className="gd-table">
              <div className="gd-tr gd-th"><span>Tab</span><span>What it is for</span></div>
              <div className="gd-tr"><span><b>Overview</b></span><span>All clients at once, worst first, with each client's score and what is pulling it down.</span></div>
              <div className="gd-tr"><span><b>Clients</b></span><span>Where the work happens. Add a client, open it, pick a month, week or day, upload reports and type in values. The client's <b>History</b> is here too.</span></div>
              <div className="gd-tr"><span><b>Formulas</b></span><span>The formula, target and source report for every KPI, by scope of work.</span></div>
              <div className="gd-tr"><span><b>Guide</b></span><span>This page.</span></div>
              <div className="gd-tr"><span><b>Q&amp;A and Feedback</b></span><span>Post a question or feedback for a department and see the answers.</span></div>
              {REMOTE ? <div className="gd-tr"><span><b>People / My team</b></span><span>People is for the Admin: managers, reporting managers, passwords and accounts. My team is for a Manager: everyone who reports to them.</span></div>
                : <div className="gd-tr"><span><b>Users</b></span><span>Admin only. Add people, reset passwords, remove accounts.</span></div>}
              <div className="gd-tr"><span><b>Audit log</b></span><span>Admin and Manager. Who uploaded or changed what, and when.{REMOTE ? " A Manager sees their own team's entries." : ""}</span></div>
            </div>
            <p>The green <b>Feedback</b> tab on the right edge of every screen is for quick notes about the app itself. You can drag it up or down if it is in your way.</p>
          </GuideSec>

          <GuideSec id="client" n="1" title="Add a client" lead="Each client is set up once, with only the scopes of work we handle for it.">
            <ol className="gd-steps">
              <li>Open the <button className="link" onClick={() => onGo("clients")}>Clients</button> tab and click <em>+ Add client</em>.</li>
              <li>Type the client's name.</li>
              <li>Tick every <b>scope of work</b> we do for this client: Account Receivable, Submission, Benefit Verification, Payment Posting, Credentialing, PCC. Account Receivable is ticked to start with. Untick it if it does not apply.</li>
              <li>Click <em>Add client</em>. The client opens straight away.</li>
            </ol>
            <GTip>Only the scopes you tick appear for that client: its upload sections, its KPIs and its SLA rows. To change them later, open the client and click <em>edit scopes</em> under its name. Switching a scope off hides it but does not erase what was saved.</GTip>
            <p><b>To remove a client</b>, click <em>Remove</em> next to it in the Clients list and type the client's name to confirm. This deletes all of its periods and cannot be undone.</p>
          </GuideSec>

          <GuideSec id="duration" n="2" title="Pick the period" lead="A period is the month, week or day the reports cover. Every period keeps its own files and its own numbers, for good.">
            <ol className="gd-steps">
              <li>Open the client. In the <b>Period</b> box click <em>Pick a month, week or day</em>.</li>
              <li>Choose a tab:
                <ul className="gd-list">
                  <li><b>Month</b>: click the month. The arrows change the year.</li>
                  <li><b>Week</b>: click any day and its whole week is picked. A week always runs <b>Monday to Sunday</b>, even when it runs into the next month, for example 28 Sep to 4 Oct.</li>
                  <li><b>Day</b>: click the day.</li>
                  <li><b>Other dates</b>: click the first day, then the last day, then <em>Use these dates</em>. Use this only for ranges that are not a month, week or day.</li>
                </ul>
              </li>
              <li>The period opens. Upload its reports and type in the rest as usual.</li>
            </ol>
            <p>Months, weeks and days that already have data carry a small dot in the picker. Picking one of them reopens it with everything in it. The four newest periods sit under the Period box as <em>Saved</em> buttons, and the rest are in the <em>All periods</em> list next to them.</p>
            <GTip>The <b>Days</b> box fills in by itself from the dates. It is used for Days in AR. Change it only if the reports cover a different number of days.</GTip>
            <GWarn>Pick weeks and months with the Week and Month tabs rather than as other dates. That way everyone's weeks start on the same Monday, and the Overview and History can line them up and compare them.</GWarn>
          </GuideSec>

          <GuideSec id="upload" n="3" title="Upload the reports" lead="There are two ways. Both give the same result, so use whichever suits you.">
            <h3>Way 1: everything at once</h3>
            {cloud ? <ol className="gd-steps">
              <li>Scroll to <b>Reports for KPI</b>. Under the OneDrive bar is the box <em>Upload this period's files from OneDrive, all at once</em>.</li>
              <li>Click <em>Upload from OneDrive</em>. OneDrive opens (see <b>Uploading from OneDrive</b> below). Open the folders down to the files, then tick the ones you want, or click <em>Upload all … files here</em>. A zip file works too.</li>
              <li>If a zip file holds more than one folder of reports, Drishti asks <em>Which folder is for this client?</em> Pick the right one.</li>
              <li>A checklist opens, headed <em>From OneDrive</em> and the folder. Every file is listed with the section it will go to. Read it.</li>
              <li>Fix any row that is not right using its dropdown, then click <em>Apply</em>. Only now are the numbers calculated.</li>
            </ol> : <ol className="gd-steps">
              <li>Scroll to <b>Reports for KPI</b>. At the top is the box <em>Drop this period's files here, all at once</em>.</li>
              <li>Drag in all the files, a whole folder, or a zip file. You can also click <em>Choose files</em> or <em>Choose a folder</em>.</li>
              <li>If what you dropped holds more than one folder of reports, Drishti asks <em>Which folder is for this client?</em> Pick the right one.</li>
              <li>A checklist opens: <em>Check these before they are applied</em>. Every file is listed with the section it will go to. Read it.</li>
              <li>Fix any row that is not right using its dropdown, then click <em>Apply</em>. Only now are the numbers calculated.</li>
            </ol>}
            <div className="gd-table">
              <div className="gd-tr gd-th"><span>Label on the row</span><span>What it means and what to do</span></div>
              <div className="gd-tr"><span><span className="chip ok">Matched</span></span><span>The file name and the columns inside both point to one section. Nothing to do.</span></div>
              <div className="gd-tr"><span><span className="chip warn">Please confirm</span></span><span>Drishti has a likely section but is not certain. Check the dropdown and change it if it is wrong.</span></div>
              <div className="gd-tr"><span><span className="chip miss">Clashes</span></span><span>Two files want the same section. Choose the section for the right file and leave the other on <em>Skip this file</em>. Apply stays switched off while two files point at one section.</span></div>
              <div className="gd-tr"><span><span className="chip info">Not recognised</span></span><span>Drishti could not tell what the file is. Pick its section from the dropdown, or leave it on <em>Skip this file</em>.</span></div>
              <div className="gd-tr"><span><span className="chip info">Not in this client</span></span><span>The file belongs to a scope this client does not have. Skip it, or add the scope with <em>edit scopes</em> and upload again.</span></div>
            </div>
            <GTip>Nothing is ever placed on a guess. If Drishti is not sure, the row starts on <em>Skip this file</em> and waits for you.</GTip>
            <h3>Way 2: one file at a time</h3>
            <ol className="gd-steps">
              <li>Under <b>Reports for KPI</b>, find the card for the report, for example <em>Denial report</em>.</li>
              {cloud ? <li>Click <em>Upload from OneDrive</em> on that card. Open the folder with the file, tick it, and click <em>Upload this file</em>.</li> : <li>Click <em>Upload file</em> on that card, or drag the file onto the card.</li>}
              <li>The card turns to <span className="tag up">loaded</span> and shows what it read: the key figures, the number of rows, the file name, and under <b>Columns read</b> the heading each figure came from.</li>
            </ol>
            <p>On a loaded card, <em>{cloud ? "Replace from OneDrive" : "Replace"}</em> swaps in a newer file and <em>Clear</em> removes it. Either way the KPIs are worked out again straight away.</p>
            <GWarn>If a figure on a card says <span style={{ color: "var(--miss)", fontWeight: 600 }}>not detected</span>, Drishti could not find that column in the file. Check the file has it, then use <em>{cloud ? "Replace from OneDrive" : "Replace"}</em>.</GWarn>
          </GuideSec>

          <GuideSec id="onedrive" title="Uploading from OneDrive" lead={`Every report file comes from the OneDrive of ${odAcct || "the team's shared Microsoft account"}. Drishti opens its My files, you pick the files, and Drishti can write a summary file back next to them. Nothing runs by itself: only when you click.`}>
            <h3>Signing in to Microsoft</h3>
            <p>The first time on a computer, clicking <em>Upload from OneDrive</em> opens a Microsoft window. Sign in as <b>{odAcct || "the shared account"}</b>, the account the whole team uses, whatever your own email is. Drishti works only with that account: if another Microsoft account is signed in, Drishti asks for this one instead. The OneDrive bar under <b>Reports for KPI</b> shows the account in use.</p>
            <h3>How the folders are kept</h3>
            <p>Keep the files in My files of that OneDrive, in a layout everyone knows. We suggest a folder called <b>Drishti</b>:</p>
            <ul className="gd-list">
              <li>In <b>Drishti</b>, one folder for each supervisor.</li>
              <li>In it, one folder for each client.</li>
              <li>In the client's folder, one folder for each month, named like <em>October 2026</em>. The month's own report files go straight in here.</li>
              <li>In the month's folder, one folder for each week, <em>Week 1</em>, <em>Week 2</em> and so on, each holding that week's report files.</li>
            </ul>
            <GTip>A week runs Monday to Friday and goes in the month its Monday is in. Week 1 is the week of the month's first Monday. In October 2026, Week 1 is 5 to 9 Oct and Week 4 is 26 to 30 Oct. Thu 1 and Fri 2 Oct belong to the week of Mon 28 Sep, so they go in September's folder. Some months have a Week 5.</GTip>
            <h3>Upload files</h3>
            <ol className="gd-steps">
              <li>Open the client and pick the month or week, as usual. The files go into the period on screen.</li>
              <li>Click <em>Upload from OneDrive</em>: in the big box under <b>Reports for KPI</b> for all of the period's files, or on one report's own card for just that report.</li>
              <li>OneDrive opens at <b>My files</b>, or where this client's last upload came from. Open the folders down to the files. The path at the top takes you back up.</li>
              <li>In the big box, tick the files you want, or click <em>Upload all … files here</em> to take every file in that folder. On a report's own card, tick the one file and click <em>Upload this file</em>.</li>
              <li>From the big box, the usual checklist opens, headed <em>From OneDrive</em> and the folder. Check it and click <em>Apply</em>. A column Drishti does not recognise brings up the same <em>Which columns?</em> pop-up as always.</li>
            </ol>
            <p><em>Pick again in OneDrive</em> in the checklist opens the folders again. Drishti only reads the report files. It never changes them. Files are not uploaded from a computer: keep every report file in OneDrive.</p>
            <h3>Update the summary</h3>
            <p>Click <em>Update Drishti summary</em>. Drishti writes <b>Drishti summary - client name.xlsx</b> in the client's folder: the one above the month and week folders you uploaded from. The bar shows where, and <em>Change</em> picks another folder. If nothing has been uploaded from OneDrive yet, Drishti asks once which folder to use. The file's <b>Monthly</b> and <b>Weekly</b> tabs list every KPI and SLA row with its formula and target, one column for each month or week coloured green, amber or red, and the change since the one before. A <b>How to read this</b> tab explains it. The file holds KPI numbers only, no patient or claim rows.</p>
            <GWarn>Drishti rewrites its Monthly and Weekly tabs each time you click Update, so keep your own notes in a tab of your own. Drishti never touches other tabs. If the file is open in the Excel desktop app, close it and click Update again.</GWarn>
            <h3>Who needs what</h3>
            <p>Everyone who uploads reports or clicks Update signs in to Microsoft with the shared account, on each computer they use, including people with a purviewservices.com email. Nothing has to be shared with anyone. Looking at clients, the Overview and History needs no Microsoft sign-in.</p>
            <GTip>If the bar says <em>OneDrive is not switched on for Drishti yet</em>, the Microsoft side has not been set up. Ask the Admin. Until then, files are dropped from the computer as before.</GTip>
          </GuideSec>

          <GuideSec id="files" title="Which files, and what they need" lead="Each scope of work has its own set of reports. A KPI can only be calculated when the reports it needs are loaded.">
            <h3>Rules for every file</h3>
            <ul className="gd-list">
              <li>Excel (.xlsx or .xls) or CSV. One report per file.</li>
              <li>Only the <b>first sheet</b> is read. Put the data there.</li>
              <li>The column headings must be within the first 15 rows, with the data directly beneath them.</li>
              <li>Column names do not have to be exact. Drishti recognises the common ones, for example Claim ID, Claim No or ICN for the claim number.</li>
            </ul>
            <h3>Naming files for the all-at-once upload</h3>
            <ul className="gd-list">
              <li>Name each file after its section, for example <em>Denial report</em> or <em>Payment Posting report</em>. Client names, months and words like "final" in the name are fine.</li>
              <li><b>Audit reports must say which department:</b> <em>AR Audit</em>, <em>Submission Audit</em> or <em>Payment Posting Audit</em>. The three look the same inside, so the name is the only way to tell them apart.</li>
              <li><b>Last period's AR inventory must say so:</b> include <em>previous</em>, <em>prior</em> or <em>last</em> in the name. Without one of those words it is treated as the current period.</li>
              <li>The payment cycle files need <em>cycle</em> or <em>paid claims</em> in the name, so they are not mistaken for the Claim Submission report.</li>
            </ul>
            <h3>The reports for each scope</h3>
            {scopesWithReports.map((sc) => (
              <div key={sc.key} className="gd-scope">
                <div className="gd-scope-h"><span className="fm-badge" style={{ background: scopeColor(sc.key) + "26", color: scopeColor(sc.key), width: 30, height: 30, borderRadius: 9, fontSize: 9.5 }}>{scopeBadge(sc.key)}</span>{sc.name}</div>
                <div className="gd-table">
                  <div className="gd-tr gd-th"><span>Report</span><span>What must be in it</span></div>
                  {[...(SCOPE_REPORTS[sc.key] || []), ...(SCOPE_SLA_REPORTS[sc.key] || [])].map((k) => <div key={k} className="gd-tr"><span><b>{REPORT_BY_KEY[k].name}</b></span><span>{REPORT_BY_KEY[k].hint}</span></div>)}
                </div>
              </div>
            ))}
            <GTip>Account Receivable also uses the <b>Payment Posting report</b>, for Collection Rate, Denial Rate and Days in AR. If the client has the Payment Posting scope, upload it there. If not, the card appears under Account Receivable.</GTip>
            <GTip>Credentialing has no report to upload. Its rows are typed in. The previous period AR inventory sits under <b>Reports for SLA</b> and is only needed for the aging comparison.</GTip>
          </GuideSec>

          <GuideSec id="columns" title="How Drishti finds the right column" lead="One report can hold several amounts side by side, for example Billed Amount, Allowed Amount and Paid Amount. Each KPI needs one particular column. This is how Drishti makes sure it reads the correct one, in every file.">
            <h3>The three checks</h3>
            <ol className="gd-steps">
              <li><b>The heading.</b> For each item it needs, Drishti looks for a heading that means that item. The headings it knows are listed below.</li>
              <li><b>The kind of data.</b> An amount must be a column of numbers. A column whose heading says date, time, ID, code, number, name, type or status is never added up as money, even when its heading shares a word with the amount. So <em>Payment Date</em> is never read as the payments, and <em>Charge Code</em> is never read as the charges.</li>
              <li><b>One clear match.</b> A specific heading beats a general one, so <em>Discrepancy Amount</em> wins over a plain <em>Amount</em> or <em>Balance</em>. Drishti uses a column only when exactly one column is a clear match.</li>
            </ol>
            <GTip>This is why a file with both <em>Paid Amount</em> and <em>Billed Amount</em> is safe. Payments are read from Paid Amount and charges from Billed Amount, because each item has its own clear column.</GTip>
            <h3>When it is not sure, it asks</h3>
            <p>Drishti stops and shows a pop-up in two situations:</p>
            <ul className="gd-list">
              <li><b>No column fits.</b> Nothing has a matching heading and the right kind of data. Example: the balance is in a column called <em>Open Dollars</em>.</li>
              <li><b>More than one column fits.</b> Example: the file has both <em>Paid Amount</em> and <em>Payment Amount</em>, or both <em>Billed Amount</em> and <em>Total Charges</em>. Drishti will not choose between them. It names the columns and asks you which one to use.</li>
            </ul>
            <p>Nothing from that report is calculated until you answer. The KPIs that depend on it stay empty rather than showing a number that might be wrong.</p>
            <h3>Your answer is remembered</h3>
            <ul className="gd-list">
              <li>The choice is saved for that client and that report, so you are not asked again next period.</li>
              <li>It is saved by the heading's name, so it still works if the columns are in a different order next time.</li>
              <li>If that heading is no longer in the file, Drishti asks again.</li>
              <li>To change a saved choice, click <em>Forget them</em> on the report's card and upload the file again.</li>
            </ul>
            <h3>How to check which column was used</h3>
            <p>Every loaded report card has a <b>Columns read</b> list. It shows each item next to the heading it was taken from, for example <em>Payments received: Paid Amount</em>. Look at it the first time you upload a client's reports. The ⓘ next to any KPI then shows the totals taken from those columns.</p>
            <GWarn>Drishti judges a column by its heading and by the kind of data in it. It cannot know that a column is mislabelled in the source report. If a heading says Paid Amount but the column holds something else, that is what will be read. A glance at <b>Columns read</b> and at the totals on the card catches this.</GWarn>
            <h3>Headings Drishti recognises</h3>
            <div className="gd-table">
              <div className="gd-tr gd-th"><span>What it needs</span><span>Headings it recognises</span></div>
              <div className="gd-tr"><span><b>Payments received</b></span><span>Payment Amount, Paid Amount, Amount Paid, Net Paid, Payment, Paid</span></div>
              <div className="gd-tr"><span><b>Charges (billed)</b></span><span>Total Charges, Charge Amount, Billed Amount, Charges, Billed</span></div>
              <div className="gd-tr"><span><b>Contractual adjustments</b></span><span>Contractual Adjustment, Contractual Write-off, Write-off, Adjustment Amount, Adjustment</span></div>
              <div className="gd-tr"><span><b>AR balance</b></span><span>AR Balance, Remaining Balance, Outstanding, Balance, Amount Due, Open Amount</span></div>
              <div className="gd-tr"><span><b>Allowed amount</b></span><span>Allowed Amount, Allowed, Allowable, Contracted Rate, Fee Schedule</span></div>
              <div className="gd-tr"><span><b>Amount received for posting</b></span><span>Received for Posting, Amount Received, Posting Amount, Remittance Amount, ERA Amount</span></div>
              <div className="gd-tr"><span><b>Discrepancy amount</b></span><span>Discrepancy, Variance, Unposted, Unable to Post, Not Posted, Difference</span></div>
              <div className="gd-tr"><span><b>Write-off (non-contractual)</b></span><span>Write-off, W/O, Adjustment Amount, with a Reason Code column so CO45 can be left out</span></div>
              <div className="gd-tr"><span><b>Recouped amount</b></span><span>Recoupment, Recouped Amount, Recovery, Take-back, Reversal, Refund</span></div>
              <div className="gd-tr"><span><b>Audit score</b></span><span>Final Audit Score, Audit Score, Final Score, Quality Score, Score</span></div>
              <div className="gd-tr"><span><b>Claim number</b></span><span>Claim ID, Claim No, Claim #, Claim Number, ICN</span></div>
              <div className="gd-tr"><span><b>Aging</b></span><span>Aging Bucket, Aging, Bucket, Age Category. If there is none, the Date of Service is used to work out the age.</span></div>
              <div className="gd-tr"><span><b>Submission date</b></span><span>Submission Date, Submitted Date, Date Submitted, Billed Date, Claim Date</span></div>
              <div className="gd-tr"><span><b>Paid date</b></span><span>Paid Date, Payment Date, Date Paid, Check Date, Remittance Date, Posted Date</span></div>
              <div className="gd-tr"><span><b>Completed or not</b></span><span>Completed, Verification Status, Status, Done, Sent, Disposition</span></div>
              <div className="gd-tr"><span><b>Urgent or regular</b></span><span>Urgent, Priority, Request Type, Verification Type</span></div>
              <div className="gd-tr"><span><b>Call outcome</b></span><span>Disposition, Call Status, Call Result, Status, Result, Outcome</span></div>
            </div>
            <p>Capital letters and spacing do not matter. <em>paid amount</em>, <em>Paid_Amount</em> and <em>PAID AMOUNT</em> are all read the same way.</p>
          </GuideSec>

          <GuideSec id="popups" title="When Drishti asks you something" lead="Drishti never guesses silently. When it is unsure it stops and asks. These are the questions you may see.">
            <h3>"Which columns?"</h3>
            <p>This appears when Drishti is not sure which column to read in a file. Either no column fits, or more than one could. The section above explains how it decides.</p>
            <ol className="gd-steps">
              <li>The pop-up names the report and the file, and lists each item it is unsure about, for example <em>Payments received</em>.</li>
              <li>Under each item it says why it is asking. If more than one column could fit, it names them.</li>
              <li>Open the dropdown and choose the column that holds that information. The likely columns are listed first, then every other column in the file.</li>
              <li>If the file truly does not have it, leave that dropdown blank. The value stays empty and you can type it in later.</li>
              <li>Click <em>Use these columns</em>. The file loads and the KPIs are calculated from the columns you chose.</li>
            </ol>
            <p>If you click <em>Cancel</em>, that file is not loaded. When several files in an all-at-once upload need this, the pop-ups come one after another.</p>
            <GTip>Your choices are remembered for that client and that report, so you answer once. To be asked again, click <em>Forget them</em> on the report's card.</GTip>
            <h3>"Confirm how to match these files"</h3>
            <p>Payment Cycle is worked out by pairing each claim in the <em>Submission file</em> with the same claim in the <em>Paid-claims file</em>.</p>
            <ul className="gd-list">
              <li>If both files have a claim identifier (Claim ID, Account No, Encounter No), the pairing is automatic and you see nothing.</li>
              <li>If they do not, Drishti proposes a combination of shared fields, such as Patient, Date of Service, CPT, Payer, Provider and Billed amount. Tick or untick the fields and approve.</li>
              <li>Your choice is remembered for this client. You are asked again only if the columns in the files change.</li>
            </ul>
            <h3>"Which folder is for this client?"</h3>
            <p>Shown when a zip file{cloud ? "" : " or folder you dropped"} holds more than one sub-folder of reports. Pick the one for the client you have open, so two clients' files are never mixed.</p>
            <h3>Status categories (Account Receivable)</h3>
            <p>The AR inventory's status column is grouped into <b>Client</b>, <b>Payer</b> and <b>Purview</b> so you can see who each balance is waiting on. Statuses Drishti cannot place show as <b>Unassigned</b>. Click <em>Edit mapping</em> in that section to assign them. The mapping is saved for that client.</p>
            <h3>"That file could not be read"</h3>
            <p>The file is damaged, password protected, or not a real Excel or CSV file. Open it, save it again as .xlsx or .csv, and upload it again.</p>
          </GuideSec>

          <GuideSec id="typed" n="4" title="Numbers you type in" lead="Some values cannot come from a report. You enter these yourself.">
            <h3>KPIs</h3>
            <ol className="gd-steps">
              <li>Under <b>KPI</b>, click a scope's bar to open it. Each row shows the metric, its value, its source and its status.</li>
              <li>The <b>Source</b> column tells you where a value came from: <span className="tag up">auto</span> was calculated from an upload, <span className="tag man">manual</span> was typed in, and <em>enter</em> means it is still empty.</li>
              <li>Fill in the empty rows. Percentages take a number such as 92. Turnaround rows take <em>Met</em>, <em>Not Met</em> or <em>N/A</em>. Some rows take Yes or No, or free text.</li>
              <li>If a value is amber or red, a box appears under it: <em>Not met, add reason</em>. Write a short reason. It is shown to leadership next to that metric.</li>
              <li>Click <em>Save</em> at the bottom of that scope. Nothing is kept until you save.</li>
            </ol>
            <GTip>You can type over an <span className="tag up">auto</span> value. It then shows as <span className="tag man">manual</span> and your typed value is the one that counts, even when more reports are uploaded later. Under the row Drishti shows what it calculated from the reports, with a link <em>Use the calculated value</em> to go back to it. Clearing your typed value and saving also brings the calculated one back.</GTip>
            <h3>SLA rows</h3>
            <ol className="gd-steps">
              <li>Scroll to <b>Service Level Agreement (SLA)</b> and open it.</li>
              <li>For each row enter the <b>Target</b>, how many were <b>Received</b> and how many were <b>Achieved</b>. Add a comment if useful.</li>
              <li>Click <em>Save</em>. Each row is scored as Achieved ÷ Received.</li>
            </ol>
            <GTip>Use the <em>Both / KPI / SLA</em> switch near the top of the client page to show only the part you are working on. Click the small ⓘ next to any metric to see its formula and the exact numbers used this period.</GTip>
          </GuideSec>

          <GuideSec id="reading" n="5" title="Read the numbers" lead="Every scored value carries one of these labels.">
            <div className="gd-table">
              <div className="gd-tr gd-th"><span>Label</span><span>Meaning</span></div>
              <div className="gd-tr"><span><span className="chip great">Excellent</span></span><span>Better than the target by a clear margin. Counts as on target.</span></div>
              <div className="gd-tr"><span><span className="chip ok">On target</span></span><span>The target is met.</span></div>
              <div className="gd-tr"><span><span className="chip warn">At risk</span></span><span>Close to the target but not meeting it. Watch it.</span></div>
              <div className="gd-tr"><span><span className="chip miss">Missed</span></span><span>The target is not met. Needs action and a reason.</span></div>
              <div className="gd-tr"><span>No label</span><span>Tracked for information only. It does not affect the score.</span></div>
            </div>
            <h3>How a client's score is worked out</h3>
            <p><span className="fm-formula">Metrics on target ÷ metrics that have a value × 100</span></p>
            <ul className="gd-list">
              <li>KPIs and SLA rows count the same, one each.</li>
              <li>Anything left blank is not counted, so an unfinished client is not marked down for missing data.</li>
              <li>The score is the share of metrics on target. It is not an average of the percentages.</li>
              <li><b>Healthy</b> at 90% or above, <b>At risk</b> from 80% to just under 90%, <b>Below target</b> under 80%.</li>
            </ul>
            <h3>A client's own page</h3>
            <p>From the Overview click <em>Open client</em>. You get the full picture for one period:</p>
            <ul className="gd-list">
              <li><b>The headline:</b> "X of Y scored metrics on target", with a bar split into on target, at risk and missed.</li>
              <li><b>A tile for every KPI:</b> the value, its label, the target, and the change from the previous period.</li>
              <li><b>The SLA table:</b> target, received, achieved and the percentage for each row.</li>
              <li><b>Account Receivable extras:</b> total inventory by aging bucket, the aging comparison against the previous period, the status wise breakdown, and who each balance is waiting on.</li>
              <li><b>Payer analysis:</b> which payers make up most of the payments, the AR and the denials.</li>
            </ul>
            <p>Use the two dropdowns at the top right of that page to change client or period.</p>
          </GuideSec>

          <GuideSec id="overview" title="The Overview" lead="One screen for leadership: every client, worst first.">
            <ol className="gd-steps">
              <li><b>Choose what to look at</b> at the top right:
                <ul className="gd-list">
                  <li><b>Latest</b>: each client's most recent period.</li>
                  <li><b>Monthly</b>: every client for the same month. Use the arrows to step back to last month and forward again.</li>
                  <li><b>Weekly</b>: the same for a week.</li>
                </ul>
                A client with nothing for the chosen month or week says so on its card.</li>
              <li><b>The top panel</b> gives the overall health across all clients, the KPI and SLA percentages, and how many clients are healthy, at risk and below target.</li>
              <li><b>Each client card</b> shows its overall score and how it moved against the period before it of the same kind: <em>vs last month</em> for a month, <em>vs last week</em> for a week. Then <em>Overall KPI</em> and <em>Overall SLA</em>, and <em>KPI by scope</em> and <em>SLA by scope</em>.</li>
              <li><b>Red 3 months in a row</b> (or weeks) names any metric that has been red in this period and the two before it. These are the problems that are not going away by themselves.</li>
              <li><b>Holding it up</b> and <b>Dragging it down</b> each show the three that matter most: the strongest, and the weakest with red before amber, taken from different scopes so you see where the problems are. A reason the supervisor typed shows under the metric. Everything else is in the small dropdown under each list, for example <em>20 more</em>, grouped by scope.</li>
              <li>Click <em>Open client</em> on a card for the detail. It opens on the month or week you were looking at.</li>
            </ol>
            <GTip>A client at 90% or above gets a <b>Star performer</b> ribbon and the gold <em>Shout-out of the period</em> banner. Click <em>Share kudos</em> to copy a ready message for your team group.</GTip>
          </GuideSec>

          <GuideSec id="history" title="History: past months and weeks" lead="Every month, week and day you upload is kept. History puts them side by side.">
            <ol className="gd-steps">
              <li>Open a client, from the Clients tab or from its card on the Overview, and click <b>History</b> at the top right.</li>
              <li>Choose <b>Monthly</b> or <b>Weekly</b> (or Daily), and how far back: the last 3 months is the default, and you can go to 6, 12 or all of them.</li>
              <li>Read across a row to see a metric month by month. Each cell is coloured green, amber or red against the target. The <b>Change</b> column compares the newest value with the one before it. For percentages the change is in points, so 64% to 70% is <em>6 pts</em>.</li>
              <li>A hatched column marked <em>Not uploaded</em> is a month or week with no reports. Upload them in that period and the gap fills in.</li>
              <li>Click a metric's name for its trend line, with the target drawn across it. Point at the line to read each value.</li>
              <li>Click a month or week heading to open that period.</li>
              <li><b>Download Excel</b> saves what is on screen, plus a second sheet with every value in rows, ready for a pivot table.</li>
            </ol>
            <div className="gd-note"><b>Example</b>Uploads start in October. In January you open the client, pick Monthly, and see October, November and December side by side, with nothing uploaded again.</div>
            {REMOTE && <p>The same numbers are also written to the Google Sheet as they are saved: the <b>KPI History</b> tab has one row per client, period and metric, and the <b>Report Details</b> tab has the totals and breakdowns from each report, including the ones the screens summarise.</p>}
            {!REMOTE && <GWarn>On this computer-only setup, history is kept in this browser. Connect the Google Sheet to keep it for everyone.</GWarn>}
          </GuideSec>

          <GuideSec id="targets" title="Targets" lead="Every KPI starts with the standard target shown on the Formulas tab.">
            <p>A client can be given its own target for any KPI. Admin and Manager can do this:</p>
            <ol className="gd-steps">
              <li>On the Overview click <em>Open client</em>.</li>
              <li>On the KPI's tile click the small gear icon.</li>
              <li>Set the direction (higher is better, lower is better, or within a band) and the values, then save.</li>
            </ol>
            <p>The new target applies to that client only. Every other client keeps the standard one.</p>
          </GuideSec>

          {REMOTE && <GuideSec id="access" title="Who sees which clients" lead="Each client is seen only by the person who maintains it, their reporting manager, and the Admin.">
            <ul className="gd-list">
              <li><b>Whoever adds a client maintains it.</b> Their name shows in the <em>Maintained by</em> column on the Clients tab and on the client's card for managers and the Admin.</li>
              <li><b>A Team Lead, SME or Supervisor</b> sees only the clients they maintain. Two team leads never see each other's numbers, even under the same manager.</li>
              <li><b>A Manager</b> sees their own clients and every client maintained by the people who report to them.</li>
              <li><b>The Admin</b> sees every client.</li>
            </ul>
            <div className="gd-note"><b>Example</b>Tina and Tom are team leads who both report to Priya. Tina adds Sunrise and Tom adds Lakeside. Tina sees only Sunrise, Tom sees only Lakeside, Priya sees both, another manager such as Rahul sees neither, and the Admin sees both.</div>
            <h3>Handing a client to someone else</h3>
            <p>A Manager or the Admin opens the Clients tab and clicks <em>Hand over</em> next to the client. A Manager can hand it to anyone in their team or keep it; the Admin can hand it to anyone. The new person sees it straight away, with every period and number.</p>
            <p>When the Admin moves someone to a different manager, their clients go with them: the new manager sees them and the old one no longer does.</p>
            <h3>Where the data lives</h3>
            <p>Everything is kept in one shared Google Sheet, so you see the same clients and numbers from any computer. Every month, week and day is kept for good, so you can look back at any of them in History. The report files themselves are read inside your browser and are never saved. Only the figures worked out from them are kept, so no patient names or claim lines are stored.</p>
            <p>When someone removes a client, it disappears from everyone's lists but its history stays. The Admin can restore it from the Clients tab.</p>
            <GWarn>If a red bar at the top says a change was not saved, check your internet connection and make the change again. Drishti also asks before you close the page while a save is still going through.</GWarn>
          </GuideSec>}

          {REMOTE ? <GuideSec id="manage" title="People, periods and the audit log">
            <h3>Managers (Admin)</h3>
            <ul className="gd-list">
              <li><b>Add a manager:</b> People tab, <em>+ Add manager</em>, enter their name and Purview email. They then create their own account with that email and choose <em>Manager</em> as their role.</li>
              <li><b>Remove a manager:</b> click <em>Remove</em> next to them. Move the people who report to them to another manager first.</li>
            </ul>
            <h3>Everyone else (Admin)</h3>
            <ul className="gd-list">
              <li><b>Change someone's reporting manager:</b> People tab, pick the new name in the <em>Reports to</em> column.</li>
              <li><b>Reset a password:</b> click <em>Reset password</em> next to the person, set a new one and share it with them directly.</li>
              <li><b>Remove an account:</b> click <em>Remove</em>. If they still maintain clients, hand those clients to someone else first. Their past entries stay in the audit log.</li>
            </ul>
            <h3>My team (Manager)</h3>
            <p>Lists everyone who reports to you, with how many clients each one maintains. To move someone or reset a password, ask the Admin.</p>
            <h3>Removed clients (Admin)</h3>
            <p>A removed client keeps all its periods. On the Clients tab click <em>Show removed clients</em>, then <em>Restore</em> to bring one back for the person who maintains it.</p>
            <h3>Removing a period (Admin)</h3>
            <p>Open the client from the Overview, choose the period at the top right, scroll to the bottom and click <em>Remove period</em>. Its reports and values are deleted, along with its rows in the KPI History and Report Details tabs. This is the only way a period is ever deleted.</p>
            <h3>Audit log (Admin and Manager)</h3>
            <p>A running record of every action: when, who, their role, what they did, the client, the period, and whether it was an upload or typed in. The Admin sees everyone's entries; a Manager sees their own and their team's.</p>
          </GuideSec> : <GuideSec id="manage" title="Users, periods and the audit log">
            <h3>Users (Admin)</h3>
            <ul className="gd-list">
              <li><b>Add a user:</b> Users tab, <em>+ Add user</em>, enter their Purview email, choose a role, set a temporary password.</li>
              <li><b>Reset a password:</b> click <em>Reset password</em> next to the person and set a new one.</li>
              <li><b>Remove a user:</b> click <em>Remove</em> and confirm. Their past entries stay in the audit log.</li>
            </ul>
            <h3>Removing a period (Admin)</h3>
            <p>Open the client from the Overview, choose the period at the top right, scroll to the bottom and click <em>Remove period</em>. Its reports and values are deleted.</p>
            <h3>Audit log (Admin and Manager)</h3>
            <p>A running record of every action: when, who, their role, what they did, the client, the period, and whether it was an upload or typed in. Use it to check who changed a number.</p>
          </GuideSec>}

          <GuideSec id="feedback" title="Questions and feedback">
            <ul className="gd-list">
              <li><b>Q&amp;A and Feedback tab:</b> choose Question or Feedback, pick the department, add the scope and client if it helps, type it and post. Answers appear under the item. Any item can be edited. Admin and Manager can delete one.</li>
              <li><b>Feedback tab on the right edge:</b> for a quick note about the app. Pick <em>Something is wrong</em>, <em>I have an idea</em> or <em>This helped me</em>, write up to 1,000 characters and send. Long numbers such as claim or account numbers are hidden automatically, so do not worry if you paste one by mistake.</li>
            </ul>
          </GuideSec>

          <GuideSec id="routine" title="Your routine each period" lead="Do this for every client once the period has closed and the reports are ready.">
            <ol className="gd-steps">
              <li>{cloud ? "Put the client's reports for the period in its folder in OneDrive, each named after its section." : "Collect the client's reports for the period and name them after their sections."}</li>
              <li>Open the client and pick the month or week with the Month or Week tab.</li>
              <li>{cloud ? <>Click <em>Upload from OneDrive</em>, pick the period's files, check the list, and click Apply.</> : <>Drop all the files in. Check the list, and click Apply.</>}</li>
              <li>Answer any pop-up, such as <em>Which columns?</em></li>
              <li>Open each scope under KPI. Fill in the rows marked <em>enter</em>, add a reason to anything amber or red, and save.</li>
              <li>Fill in the SLA rows and save.</li>
              {cloud && <li>Click <em>Update Drishti summary</em>.</li>}
              <li>Open the Overview, switch to Monthly or Weekly, and check the client's score reads as you expect. History shows how it compares with earlier months.</li>
            </ol>
          </GuideSec>

          <GuideSec id="help" title="If something looks wrong">
            <div className="gd-table">
              <div className="gd-tr gd-th"><span>What you see</span><span>What to do</span></div>
              <div className="gd-tr"><span>A KPI is empty after uploading</span><span>One of the reports it needs is missing. Check the Formulas tab for which reports feed it, and look for a card that is not marked loaded.</span></div>
              <div className="gd-tr"><span>A number looks wrong</span><span>Click the ⓘ next to it. It shows the formula and the exact figures taken from your files. If a figure is off, check <b>Columns read</b> on that report's card. If the wrong column was used, click <em>Forget them</em> if it was a saved choice, fix the heading in the report if needed, and upload again.</span></div>
              <div className="gd-tr"><span>A file says Not recognised</span><span>Rename it after its section, or just pick the section from the dropdown in the checklist.</span></div>
              <div className="gd-tr"><span>A client shows nothing for the chosen month or week</span><span>Nothing was uploaded for that client in that month or week, or it was saved with Other dates instead of the Month or Week tab. Open the client, pick the month or week with its own tab and upload there.</span></div>
              <div className="gd-tr"><span>The score is lower than expected</span><span>The score counts how many metrics are on target. One amber or red SLA row counts the same as one KPI. Check <em>Dragging it down</em> on the client's card.</span></div>
              <div className="gd-tr"><span>Payment Cycle is blank</span><span>Both the Submission file and the Paid-claims file are needed, each with a date column, and Drishti may be waiting for you to confirm how to match them.</span></div>
              {local && <div className="gd-tr"><span>Drishti is empty on another computer</span><span>Your work is saved in the browser where you entered it. Go back to that computer and browser.</span></div>}
              {REMOTE && <div className="gd-tr"><span>A client you expect is not in your list</span><span>You only see clients you maintain (and your team's, if you are a Manager). Ask your manager or the Admin to hand it to you.</span></div>}
              {REMOTE && <div className="gd-tr"><span>Your manager is not in the list when creating an account</span><span>The Admin has not added them yet. Ask the Admin to add your manager on the People tab.</span></div>}
              {REMOTE && <div className="gd-tr"><span>Payment Cycle did not change after uploading one file again</span><span>Payment Cycle matches the Submission file against the Paid-claims file at the moment you upload them. Upload both files together again to recalculate it.</span></div>}
              <div className="gd-tr"><span>You cannot see {REMOTE ? "People, My team" : "Users"} or Audit log</span><span>Your role does not include them. Ask {REMOTE ? "the Admin" : "an Admin"}.</span></div>
            </div>
            <p>Still stuck? Use the Feedback tab on the right edge and choose <em>Something is wrong</em>.</p>
          </GuideSec>
        </div>
      </div>
    </div>
  );
}

/* =============================== users ================================ */
function Users({ me, users, refreshUsers }) {
  const [open, setOpen] = useState(false); const [resetFor, setResetFor] = useState(null);
  const [delFor, setDelFor] = useState(null); const [notice, setNotice] = useState("");
  const addUser = async ({ username, role, password }) => {
    if (users.some((u) => u.username.toLowerCase() === username.trim().toLowerCase())) { setNotice("An account with that email already exists."); setOpen(false); return; }
    setNotice("");
    const salt = uid(); const u = { id: uid(), username: username.trim(), role: isNamedAdmin(username) ? "admin" : role === "admin" ? "supervisor" : role, salt, passHash: await sha256(salt + password), createdAt: nowIso(), createdBy: me.username };
    await sSet(K.users, [...users, u]); await refreshUsers(); await appendAudit({ user: me.username, role: me.role, action: `created ${role} “${u.username}”` }); setOpen(false);
  };
  const resetPass = async (t, np) => { const salt = uid(); const nu = await Promise.all(users.map(async (x) => x.id === t.id ? { ...x, salt, passHash: await sha256(salt + np) } : x)); await sSet(K.users, nu); await refreshUsers(); await appendAudit({ user: me.username, role: me.role, action: `reset password for “${t.username}”` }); setResetFor(null); };
  const askDelete = (u) => { if (u.id === me.id) { setNotice("You cannot remove your own account while you are signed in."); return; } setNotice(""); setDelFor(u); };
  const delUser = async () => { const u = delFor; if (!u) return; setDelFor(null); await sSet(K.users, users.filter((x) => x.id !== u.id)); await refreshUsers(); await appendAudit({ user: me.username, role: me.role, action: `removed user “${u.username}”` }); };
  return (
    <div className="wrap">
      <div className="pagehead"><div><h1>Users</h1><div className="sub">{users.length} account{users.length !== 1 ? "s" : ""}</div></div><button className="btn primary" onClick={() => setOpen(true)}>+ Add user</button></div>
      {notice && <div className="rc-err" style={{ marginBottom: 10 }}>{notice}</div>}
      <div className="note-box">Passwords are scrambled before they are saved. This sign-in suits an internal, trusted team. It is not a substitute for a full server-side login if outside users or sensitive data are ever involved.</div>
      <div className="panel"><table><thead><tr><th>Email</th><th>Role</th><th>Created</th><th></th></tr></thead>
        <tbody>{users.map((u) => (<tr key={u.id}><td style={{ color: "var(--ink)", fontWeight: 500 }}>{u.username}{u.id === me.id && <span style={{ color: "var(--ink-3)", fontWeight: 400 }}> · you</span>}</td><td><span className="rolechip">{u.role}</span></td><td className="mono" style={{ fontSize: 12 }}>{fmtDateTime(u.createdAt)}</td><td style={{ textAlign: "right", whiteSpace: "nowrap" }}><button className="btn sm" onClick={() => setResetFor(u)}>Reset password</button>{" "}<button className="btn sm ghost" onClick={() => askDelete(u)}>Remove</button></td></tr>))}</tbody>
      </table></div>
      {open && <AddUser onClose={() => setOpen(false)} onCreate={addUser} />}
      {resetFor && <AdminReset user={resetFor} onClose={() => setResetFor(null)} onReset={resetPass} />}
      {delFor && <Modal title="Remove user" onClose={() => setDelFor(null)}>
        <p style={{ fontSize: 13, color: "var(--ink-2)", marginTop: 0 }}>Remove <b>{delFor.username}</b>? They will lose access straight away. Their entries in the audit log are kept.</p>
        <div className="modal-actions"><button className="btn ghost" onClick={() => setDelFor(null)}>Cancel</button><button className="btn danger solid" onClick={delUser}>Remove user</button></div>
      </Modal>}
    </div>
  );
}

/* ============== people on the shared server: Admin sees all, a Manager sees their team ============== */
function People({ me, users, clients, refreshUsers }) {
  const isAdmin = me.role === "admin";
  const [mgrs, setMgrs] = useState(null);
  const [notice, setNotice] = useState(""); const [okNote, setOkNote] = useState("");
  const [addMgr, setAddMgr] = useState(false); const [delMgr, setDelMgr] = useState(null);
  const [resetFor, setResetFor] = useState(null); const [delFor, setDelFor] = useState(null);
  const loadMgrs = async () => { if (!isAdmin) return; try { setMgrs((await api("managers.list")).managers || []); } catch (e) { setNotice(e.message); setMgrs([]); } };
  useEffect(() => { loadMgrs(); refreshUsers(); }, []);
  const run = async (fn, ok) => { setNotice(""); setOkNote(""); try { await fn(); if (ok) setOkNote(ok); } catch (e) { setNotice(e.message); } await Promise.all([refreshUsers(), loadMgrs()]); };
  const owned = (id) => (clients || []).filter((c) => c.ownerId === id).length;
  const people = users.filter((u) => u.id !== "admin");
  const team = isAdmin ? people : people.filter((u) => u.id !== me.id);
  return (
    <div className="wrap">
      <div className="pagehead"><div><h1>{isAdmin ? "People" : "My team"}</h1><div className="sub">{isAdmin ? "Managers you set up, and everyone who has created an account" : "Everyone who reports to you. You see every client they maintain."}</div></div>
        {isAdmin && <button className="btn primary" onClick={() => setAddMgr(true)}>+ Add manager</button>}</div>
      {notice && <div className="rc-err" style={{ marginBottom: 10 }}>{notice}</div>}
      {okNote && <div className="msg ok">{okNote}</div>}
      {isAdmin && <>
        <h3 className="people-h">Managers</h3>
        <div className="note-box">Everyone who creates an account must choose a reporting manager from this list. A manager creates their own account with the email you enter here, choosing <b>Manager</b> as their role, and then reports to you.</div>
        {mgrs == null ? <div className="card" style={{ color: "var(--ink-3)" }}>Loading…</div> : mgrs.length === 0 ? <div className="empty"><p>No managers yet. Add your first manager so that supervisors, TLs and SMEs can pick them when they create an account.</p></div> :
          <div className="panel"><table><thead><tr><th>Manager</th><th>Email</th><th>Account</th><th>People reporting</th><th></th></tr></thead>
            <tbody>{mgrs.map((m) => (<tr key={m.id}><td style={{ color: "var(--ink)", fontWeight: 600 }}>{m.name}</td><td>{m.email}</td><td>{m.signedUp ? <span className="chip ok">Created</span> : <span className="chip info">Not yet</span>}</td><td className="mono">{m.teamSize}</td><td style={{ textAlign: "right" }}><button className="btn sm ghost" onClick={() => setDelMgr(m)}>Remove</button></td></tr>))}</tbody>
          </table></div>}
        <h3 className="people-h">Everyone with an account</h3>
      </>}
      {team.length === 0 ? <div className="empty"><p>{isAdmin ? "Nobody has created an account yet." : "Nobody reports to you yet. People appear here once they create an account and choose you as their reporting manager."}</p></div> :
        <div className="panel"><table><thead><tr><th>Name</th><th>Email</th><th>Role</th>{isAdmin && <th>Reports to</th>}<th>Clients</th><th>Joined</th>{isAdmin && <th></th>}</tr></thead>
          <tbody>{team.map((u) => (<tr key={u.id}>
            <td style={{ color: "var(--ink)", fontWeight: 600 }}>{u.name || "—"}{u.id === me.id && <span style={{ color: "var(--ink-3)", fontWeight: 400 }}> · you</span>}</td>
            <td>{u.username}</td>
            <td><span className="rolechip">{roleLabel(u.role)}</span></td>
            {isAdmin && <td>{u.role === "manager" ? <span style={{ color: "var(--ink-3)" }}>{me.name || "Admin"}</span> :
              <select className="inline-select" value={u.managerId || ""} onChange={(e) => run(() => api("users.setManager", { userId: u.id, managerId: e.target.value }), (u.name || u.username) + " now reports to " + ((mgrs || []).find((m) => m.id === e.target.value)?.name || "the new manager") + ".")}>
                {!u.managerId && <option value="">Choose…</option>}{(mgrs || []).map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
              </select>}</td>}
            <td className="mono">{owned(u.id)}</td>
            <td className="mono" style={{ fontSize: 12 }}>{u.createdAt ? fmtDateTime(u.createdAt) : "—"}</td>
            {isAdmin && <td style={{ textAlign: "right", whiteSpace: "nowrap" }}><button className="btn sm" onClick={() => setResetFor(u)}>Reset password</button>{" "}<button className="btn sm ghost" onClick={() => setDelFor(u)}>Remove</button></td>}
          </tr>))}</tbody>
        </table></div>}
      {!isAdmin && <div className="note-box" style={{ marginTop: 14 }}>To move someone to a different manager, or to reset a password, ask the Admin.</div>}
      {addMgr && <AddManager onClose={() => setAddMgr(false)} onSave={async (name, email) => { setAddMgr(false); await run(() => api("managers.add", { name, email }), name + " is now on the managers list. Ask them to create their account with " + email + ", choosing Manager as their role."); }} />}
      {delMgr && <Modal title={"Remove " + delMgr.name + "?"} onClose={() => setDelMgr(null)}>
        <p style={{ fontSize: 13, color: "var(--ink-2)", marginTop: 0 }}>{delMgr.teamSize ? <>{delMgr.teamSize} {delMgr.teamSize > 1 ? "people report" : "person reports"} to {delMgr.name}. Move them to another manager first, using the Reports to column.</> : <>{delMgr.name} comes off the managers list. If they have an account, it becomes a Supervisor account.</>}</p>
        <div className="modal-actions"><button className="btn ghost" onClick={() => setDelMgr(null)}>Cancel</button>{!delMgr.teamSize && <button className="btn danger solid" onClick={async () => { const m = delMgr; setDelMgr(null); await run(() => api("managers.remove", { id: m.id }), m.name + " was removed from the managers list."); }}>Remove manager</button>}</div>
      </Modal>}
      {resetFor && <AdminReset user={resetFor} onClose={() => setResetFor(null)} onReset={async (u, pw) => { setResetFor(null); await run(() => api("users.resetPassword", { userId: u.id, password: pw }), "New password set for " + (u.name || u.username) + ". Share it with them directly."); }} />}
      {delFor && <Modal title={"Remove " + (delFor.name || delFor.username) + "?"} onClose={() => setDelFor(null)}>
        <p style={{ fontSize: 13, color: "var(--ink-2)", marginTop: 0 }}>{owned(delFor.id) ? <>{delFor.name || delFor.username} still maintains {owned(delFor.id)} client{owned(delFor.id) > 1 ? "s" : ""}. Hand {owned(delFor.id) > 1 ? "them" : "it"} to someone else from the Clients tab first.</> : <>They lose access straight away. Their entries in the audit log are kept.</>}</p>
        <div className="modal-actions"><button className="btn ghost" onClick={() => setDelFor(null)}>Cancel</button>{!owned(delFor.id) && <button className="btn danger solid" onClick={async () => { const u = delFor; setDelFor(null); await run(() => api("users.remove", { userId: u.id }), (u.name || u.username) + " was removed."); }}>Remove account</button>}</div>
      </Modal>}
    </div>
  );
}
function AddManager({ onClose, onSave }) {
  const [name, setName] = useState(""); const [email, setEmail] = useState("");
  const ok = name.trim() && EMAIL_RX.test(email.trim());
  return (<Modal onClose={onClose} title="Add manager" desc="Supervisors, TLs and SMEs will see this name when they choose their reporting manager.">
    <div className="field"><label>Manager's name</label><input autoFocus value={name} onChange={(e) => setName(e.target.value)} placeholder="First and last name" /></div>
    <div className="field"><label>Manager's Purview email</label><input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="name@purview-healthcare.com" /></div>
    <Actions onClose={onClose} onOk={() => onSave(name.trim(), email.trim().toLowerCase())} okLabel="Add manager" okDisabled={!ok} />
  </Modal>);
}

/* ============================= audit log ============================== */
const DEPARTMENTS = ["Account Receivable", "Submission", "Payment Posting", "Benefit Verification", "Credentialing", "PCC", "General", "Application / Tool"];
const SCOPE_OPTIONS = ["Account Receivable", "Submission", "Payment Posting", "Benefit Verification", "Credentialing", "PCC"];
function Questions({ me, clients }) {
  const [items, setItems] = useState(null);
  const [tab, setTab] = useState("question");
  const [dept, setDept] = useState("Submission");
  const [scope, setScope] = useState("");
  const [client, setClient] = useState("");
  const [text, setText] = useState("");
  const [answers, setAnswers] = useState({});
  const [delId, setDelId] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [editText, setEditText] = useState("");
  const canAdd = ["admin", "manager", "supervisor", "tl", "sme"].includes(me.role);
  const canDelete = ["admin", "manager"].includes(me.role);
  const fmtDate = (x) => { try { return new Date(x).toLocaleDateString(undefined, { month: "short", day: "numeric" }) + " " + new Date(x).toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" }); } catch { return x; } };
  useEffect(() => { (async () => {
    const SEED = [
      { id: "seed-sub-resub", kind: "question", dept: "Submission", scope: "Submission", question: "Does the submission team handle any claim resubmissions? If yes, please specify for which clients." },
      { id: "seed-ar-samples", kind: "task", dept: "Account Receivable", scope: "Account Receivable", question: "Please share at least 2 sample AR-scope reports from 2 different clients (dummy patient data is fine): AR Inventory, Denial, Payment Posting, and AR Audit reports, plus the Escalation Tracker, so we can validate the automatic column detection and KPI calculations against real column names." },
      { id: "seed-sub-samples", kind: "task", dept: "Submission", scope: "Submission", question: "Please share at least 2 sample Submission-scope reports from 2 different clients (dummy data is fine): Claim Submission, Claim Rejection, Submission Inflow, and Submission Audit reports, so we can validate the auto-calculations (Clean Claim, Rejection Rate, Inflow vs Submitted, Quality) against real column names." },
      { id: "seed-pp-samples", kind: "task", dept: "Payment Posting", scope: "Payment Posting", question: "Please share at least 2 sample Payment Posting / ERA reports from 2 different clients (dummy data is fine) with columns such as payer, payer type, billed, allowed, paid, contractual adjustment, recoupment, and the relevant dates, so we can set up and validate the Payment Posting metrics." },
      { id: "seed-ev-tat", kind: "question", dept: "Benefit Verification", scope: "Benefit Verification", question: "For the EV TAT metrics (New Patients TAT – Same Day, Urgent Request TAT – 3 hr, TAT – 48–72 hr): should each be auto-calculated as a % within target using timestamps from the uploaded EV report, or should the supervisor simply mark Met / Not Met manually?\n\nWhy auto MIGHT work: if the EV report has a 'received date/time' and a 'completed date/time' for every request, the tool can compute turnaround per request and report, e.g. '% of urgent requests completed within 3 hrs', which is accurate and needs no manual judgement.\n\nWhy auto MIGHT NOT work: if the report doesn't capture both timestamps (e.g. only a completed date, or verifications done on payer portals with no logged received time), there is nothing to calculate from, so the supervisor would mark Met / Not Met from their own tracking.\n\nPlease confirm for each TAT: auto from timestamps or manual Met/Not-Met, and whether the EV report contains both received and completed date/time columns." },
      { id: "seed-ev-samples", kind: "task", dept: "Benefit Verification", scope: "Benefit Verification", question: "Please share at least 2 sample EV / Benefit Verification reports from 2 different clients (dummy patient data is fine) so we can see the report format and set up the EV metric calculations: TATs, Urgent vs Regular verification, and Quality. Specifically we need to see whether the report includes: a Request Type column (Urgent / Regular / New Patient), and received + completed date/time columns." },
      { id: "seed-pcc-samples", kind: "task", dept: "PCC", scope: "PCC", question: "Please share at least 2 sample PCC reports from 2 different clients (dummy data is fine) so we can see the format and set up the PCC metrics. Also please list which metrics your team tracks and the formula + standard for each." },
    ];
    let q = await sGet(K.questions);
    const firstTime = q == null;
    q = q || [];
    let seeded = await sGet(K.qseed) || (firstTime ? [] : ["seed-sub-resub"]);
    let changed = false;
    for (const s of SEED) {
      const existing = q.find((x) => x.id === s.id);
      if (existing) { if (!existing.kind) { existing.kind = s.kind; changed = true; } if (existing.status !== "answered" && !existing.editedBy && existing.question !== s.question) { existing.question = s.question; changed = true; } }
      else if (!seeded.includes(s.id)) { q = [{ id: s.id, kind: s.kind, dept: s.dept, scope: s.scope, client: "", question: s.question, answer: "", status: "open", askedBy: "system", askedAt: nowIso() }, ...q]; seeded = [...seeded, s.id]; changed = true; }
    }
    if (changed) { await sSet(K.questions, q); await sSet(K.qseed, seeded); }
    setItems(q);
  })(); }, []);
  // each change is applied to the newest list, so two people posting at once do not wipe out each other's posts
  const save = async (change) => { const base = REMOTE ? ((await sGet(K.questions)) || items) : items; const next = change(base); setItems(next); await sSet(K.questions, next); };
  const add = async () => { if (!text.trim()) return; const item = { id: uid(), kind: tab, dept, scope: scope || "", client: client || "", question: text.trim(), answer: "", status: "open", askedBy: me.username, askedAt: nowIso() }; setText(""); setScope(""); setClient(""); await save((list) => [item, ...list]); Promise.resolve(appendAudit({ user: me.username, role: me.role, action: "posted " + (tab === "task" ? "feedback" : "question"), client: dept })).catch(() => {}); };
  const answer = async (id) => { const a = answers[id]; if (a == null || !a.trim()) return; await save((list) => list.map((q) => q.id === id ? { ...q, answer: a.trim(), status: "answered", answeredBy: me.username, answeredAt: nowIso() } : q)); };
  const reopen = async (id) => save((list) => list.map((q) => q.id === id ? { ...q, status: "open" } : q));
  const startEdit = (q) => { setEditingId(q.id); setEditText(q.question); };
  const saveEdit = async () => { if (!editText.trim()) return; const id = editingId; setEditingId(null); const txt = editText.trim(); await save((list) => list.map((q) => q.id === id ? { ...q, question: txt, editedBy: me.username, editedAt: nowIso() } : q)); };
  const del = (id) => setDelId(id);
  const doDel = () => { const id = delId; save((list) => list.filter((q) => q.id !== id)); setDelId(null); };
  if (items == null) return <div className="wrap"><div className="card" style={{ color: "var(--ink-3)" }}>Loading…</div></div>;
  const isTask = tab === "task";
  const view = items.filter((q) => (q.kind || "question") === tab);
  const byDept = DEPARTMENTS.map((d) => [d, view.filter((q) => q.dept === d)]).filter(([, qs]) => qs.length);
  const openCount = view.filter((q) => q.status !== "answered").length;
  const metaLine = (q) => ["asked by " + q.askedBy, q.scope && q.scope !== q.dept ? "scope: " + q.scope : null, q.client ? "client: " + q.client : null, fmtDate(q.askedAt), q.editedAt ? "edited" : null].filter(Boolean).join(" · ");
  return (
    <div className="wrap">
      <div className="pagehead"><div><h1>Q&amp;A and Feedback</h1><div className="sub">Post questions for departments or share feedback · answered on each team's own shift</div></div></div>
      <div className="seg" style={{ marginBottom: 16 }}>
        <button className={"segbtn" + (tab === "question" ? " on" : "")} onClick={() => setTab("question")}>Questions</button>
        <button className={"segbtn" + (tab === "task" ? " on" : "")} onClick={() => setTab("task")}>Feedback</button>
      </div>
      {canAdd && <div className="card" style={{ marginBottom: 16 }}>
        <div className="row" style={{ gap: 8, flexWrap: "wrap", marginBottom: 8 }}>
          <div style={{ flex: "1 1 150px" }}><label className="flab">Department</label><select value={dept} onChange={(e) => setDept(e.target.value)} style={{ width: "100%" }}>{DEPARTMENTS.map((d) => <option key={d}>{d}</option>)}</select></div>
          <div style={{ flex: "1 1 150px" }}><label className="flab">Scope of work</label><select value={scope} onChange={(e) => setScope(e.target.value)} style={{ width: "100%" }}><option value="">— optional —</option>{SCOPE_OPTIONS.map((d) => <option key={d}>{d}</option>)}</select></div>
          <div style={{ flex: "1 1 150px" }}><label className="flab">Client</label><select value={client} onChange={(e) => setClient(e.target.value)} style={{ width: "100%" }}><option value="">— optional / all —</option>{(clients || []).map((c) => <option key={c.id}>{c.name}</option>)}</select></div>
        </div>
        <div className="row" style={{ gap: 10, alignItems: "flex-start" }}>
          <textarea value={text} onChange={(e) => setText(e.target.value)} placeholder={isTask ? "Share your feedback…" : "Type your question…"} style={{ flex: 1, minHeight: 44, resize: "vertical" }} />
          <button className="btn primary" onClick={add}>Post {isTask ? "feedback" : "question"}</button>
        </div>
      </div>}
      <div style={{ fontSize: 12, color: "var(--ink-3)", marginBottom: 10 }}>{openCount} open {isTask ? "feedback item" : "question"}{openCount !== 1 ? "s" : ""}</div>
      {byDept.length === 0 && <div className="empty"><p>No {isTask ? "feedback" : "questions"} yet.</p></div>}
      {byDept.map(([d, qs]) => (
        <div key={d} style={{ marginBottom: 18 }}>
          <div className="section-h"><h2>{d} <span style={{ color: "var(--ink-3)", fontWeight: 400, fontSize: 13 }}>· {qs.length}</span></h2><span className="rule" /></div>
          {qs.map((q) => (
            <div key={q.id} className="card" style={{ marginBottom: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start" }}>
                {editingId === q.id
                  ? <textarea value={editText} onChange={(e) => setEditText(e.target.value)} style={{ flex: 1, minHeight: 46, resize: "vertical" }} />
                  : <div style={{ fontSize: 13.5, fontWeight: 500, whiteSpace: "pre-wrap", lineHeight: 1.5 }}>{q.question}</div>}
                <span className={"chip " + (q.status === "answered" ? "ok" : "warn")} style={{ whiteSpace: "nowrap" }}>{q.status === "answered" ? (isTask ? "Done" : "Answered") : "Open"}</span>
              </div>
              {editingId === q.id
                ? <div className="row" style={{ justifyContent: "flex-end", marginTop: 8, gap: 8 }}><button className="btn sm ghost" onClick={() => setEditingId(null)}>Cancel</button><button className="btn sm primary" onClick={saveEdit}>Save</button></div>
                : <div style={{ fontSize: 11, color: "var(--ink-3)", marginTop: 4 }}>{metaLine(q)}</div>}
              {q.status === "answered"
                ? <div style={{ marginTop: 10, background: "var(--ok-bg)", borderRadius: 8, padding: "9px 11px", fontSize: 12.5 }}><b>{isTask ? "Response:" : "Answer:"}</b> {q.answer}<div style={{ fontSize: 11, color: "var(--ink-3)", marginTop: 4 }}>by {q.answeredBy} · {fmtDate(q.answeredAt)}</div></div>
                : <div style={{ marginTop: 10 }}><textarea placeholder={isTask ? "Add a response / mark done…" : "Type the answer…"} value={answers[q.id] ?? ""} onChange={(e) => setAnswers((s) => ({ ...s, [q.id]: e.target.value }))} style={{ width: "100%", minHeight: 40, resize: "vertical" }} /><div className="row" style={{ justifyContent: "flex-end", marginTop: 6 }}><button className="btn sm primary" onClick={() => answer(q.id)}>{isTask ? "Mark done" : "Save answer"}</button></div></div>}
              {(canAdd || q.status === "answered" || canDelete) && editingId !== q.id && <div className="row" style={{ justifyContent: "flex-end", marginTop: 6, gap: 8 }}>{canAdd && <button className="btn sm ghost" onClick={() => startEdit(q)}>Edit</button>}{q.status === "answered" && <button className="btn sm ghost" onClick={() => reopen(q.id)}>Reopen</button>}{canDelete && <button className="btn sm ghost" onClick={() => del(q.id)}>Delete</button>}</div>}
            </div>
          ))}
        </div>
      ))}
      {delId && <Modal title="Delete this item" onClose={() => setDelId(null)}>
        <p style={{ fontSize: 13, color: "var(--ink-2)", marginTop: 0 }}>This removes the entry and any answer recorded against it. This cannot be undone.</p>
        <div className="modal-actions"><button className="btn ghost" onClick={() => setDelId(null)}>Cancel</button><button className="btn danger solid" onClick={doDel}>Delete</button></div>
      </Modal>}
    </div>
  );
}

function AuditLog() {
  const [log, setLog] = useState(null);
  useEffect(() => { (async () => setLog((await sGet(K.audit)) || []))(); }, []);
  if (log === null) return <div className="wrap"><div className="card" style={{ color: "var(--ink-3)" }}>Loading…</div></div>;
  return (
    <div className="wrap">
      <div className="pagehead"><div><h1>Audit log</h1><div className="sub">{log.length} recent event{log.length !== 1 ? "s" : ""}</div></div></div>
      {log.length === 0 ? <div className="empty"><p>No activity recorded yet.</p></div> :
        <div className="panel"><table><thead><tr><th>When</th><th>User</th><th>Role</th><th>Action</th><th>Client</th><th>Period</th><th>Source</th></tr></thead>
          <tbody>{log.map((e) => (<tr key={e.id}><td className="mono" style={{ fontSize: 12, whiteSpace: "nowrap" }}>{fmtDateTime(e.ts)}</td><td style={{ color: "var(--ink)", fontWeight: 500 }}>{e.user}</td><td><span className="rolechip">{roleLabel(e.role)}</span></td><td>{e.action}</td><td>{e.client || "—"}</td><td>{e.period || "—"}</td><td>{e.method ? <span className="tag">{e.method}{e.file ? " · " + e.file : ""}</span> : "—"}</td></tr>))}</tbody>
        </table></div>}
    </div>
  );
}

/* ============================== modals ================================ */
function ScopeChecklist({ value, onChange }) {
  const toggle = (k) => onChange(value.includes(k) ? value.filter((x) => x !== k) : [...value, k]);
  return (<div className="scopepick">{SCOPES.map((s) => (
    <label key={s.key} className={value.includes(s.key) ? "on" : ""}><input type="checkbox" checked={value.includes(s.key)} onChange={() => toggle(s.key)} /><div><div style={{ fontWeight: 500 }}>{s.name} <span style={{ color: "var(--ink-3)", fontWeight: 400 }}>· {s.short}</span></div><div className="d">{s.criteria ? "KPI + SLA" : "SLA only"}</div></div></label>
  ))}</div>);
}
function AddClient({ onClose, onCreate }) {
  const [name, setName] = useState(""); const [scopes, setScopes] = useState(["ar"]);
  return (<Modal onClose={onClose} title="Add client" desc="Name the client and pick the scopes of work they deal with. Only those will show up in this client's profile. You can add more later.">
    <div className="field"><label>Client name</label><input autoFocus value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. PHC" /></div>
    <div className="field"><label>Scopes of work</label><ScopeChecklist value={scopes} onChange={setScopes} /></div>
    <Actions onClose={onClose} onOk={() => onCreate({ name, scopes })} okLabel="Add client" okDisabled={!name.trim() || scopes.length === 0} />
  </Modal>);
}
function EditScopes({ client, onClose, onSave }) {
  const [scopes, setScopes] = useState(client.scopes || []);
  return (<Modal onClose={onClose} title={`Scopes: ${client.name}`} desc="Turn scopes on or off for this client. Removing a scope hides its fields but doesn't erase saved numbers.">
    <ScopeChecklist value={scopes} onChange={setScopes} />
    <Actions onClose={onClose} onOk={() => onSave(scopes)} okLabel="Save scopes" okDisabled={scopes.length === 0} />
  </Modal>);
}
function AddScope({ current, onClose, onAdd }) {
  const avail = SCOPES.filter((s) => !current.includes(s.key));
  return (<Modal onClose={onClose} title="Add a scope of work" desc="Add another scope to this client when they take on new work.">
    {avail.length === 0 ? <p>All scopes are already active for this client.</p> :
      <div className="scopepick">{avail.map((s) => <label key={s.key} onClick={() => onAdd(s.key)} style={{ cursor: "pointer" }}><div><div style={{ fontWeight: 500 }}>{s.name} <span style={{ color: "var(--ink-3)", fontWeight: 400 }}>· {s.short}</span></div><div className="d">{s.kind === "kpi" ? s.criteria.length + " criteria" : s.tasks.length + " tasks"}</div></div></label>)}</div>}
    <div className="modal-actions"><button className="btn ghost" onClick={onClose}>Close</button></div>
  </Modal>);
}
function NewPeriod({ onClose, onCreate }) {
  const [kind, setKind] = useState("week"); const [label, setLabel] = useState(""); const [from, setFrom] = useState(""); const [to, setTo] = useState("");
  const autoLabel = () => {
    if (label.trim()) return label.trim();
    if (from.trim()) return to.trim() ? `${from.trim()} – ${to.trim()}` : from.trim();
    return `${kind === "week" ? "Week" : kind === "month" ? "Month" : "Period"} · ${new Date().toLocaleDateString()}`;
  };
  return (<Modal onClose={onClose} title="New period" desc="Numbers are stored under this period. Files are not kept.">
    <div className="field"><label>Type</label><select value={kind} onChange={(e) => setKind(e.target.value)}><option value="week">Weekly</option><option value="month">Monthly</option><option value="custom">Custom</option></select></div>
    <div className="field"><label>Label <span style={{ color: "var(--ink-3)", fontWeight: 400 }}>(optional, auto-named from the dates if left blank)</span></label><input autoFocus value={label} onChange={(e) => setLabel(e.target.value)} placeholder={kind === "week" ? "e.g. Week 1 – April" : kind === "month" ? "e.g. April 2025" : "e.g. 04/01 – 04/20"} /></div>
    <div className="two"><div className="field"><label>From (optional)</label><input value={from} onChange={(e) => setFrom(e.target.value)} placeholder="04/01" /></div><div className="field"><label>To (optional)</label><input value={to} onChange={(e) => setTo(e.target.value)} placeholder="04/07" /></div></div>
    <Actions onClose={onClose} onOk={() => onCreate({ label: autoLabel(), kind, from, to })} okLabel="Create period" />
  </Modal>);
}
function AddUser({ onClose, onCreate }) {
  const [username, setUsername] = useState(""); const [role, setRole] = useState("supervisor"); const [password, setPassword] = useState("");
  const email = username.trim().toLowerCase();
  const okEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  const ok = okEmail && password.length >= 6;
  return (<Modal onClose={onClose} title="Add user" desc="They sign in with this email and password. You can reset their password from this page at any time.">
    <div className="two"><div className="field"><label>Purview email</label><input autoFocus type="email" value={username} onChange={(e) => setUsername(e.target.value)} placeholder="name@purview-healthcare.com" /></div><div className="field"><label>Role</label><select value={role} onChange={(e) => setRole(e.target.value)}><option value="supervisor">Supervisor</option><option value="manager">Manager</option></select></div></div>
    <div className="field"><label>Temporary password (at least 6 characters)</label><input type="password" value={password} onChange={(e) => setPassword(e.target.value)} /></div>
    <Actions onClose={onClose} onOk={() => onCreate({ username: email, role, password })} okLabel="Create user" okDisabled={!ok} />
  </Modal>);
}
function AdminReset({ user, onClose, onReset }) {
  const [pw, setPw] = useState("");
  return (<Modal onClose={onClose} title={`Reset password: ${user.username}`} desc="Set a new password for this user.">
    <div className="field"><label>New password</label><input autoFocus type="password" value={pw} onChange={(e) => setPw(e.target.value)} /></div>
    <Actions onClose={onClose} onOk={() => onReset(user, pw)} okLabel="Reset password" okDisabled={!pw} />
  </Modal>);
}
function EditRule({ scopeName, crit, rule, onClose, onSave }) {
  const [dir, setDir] = useState(rule.dir); const [fmt, setFmt] = useState(rule.fmt);
  const [good, setGood] = useState(rule.good ?? ""); const [fair, setFair] = useState(rule.fair ?? "");
  const [min, setMin] = useState(rule.min ?? ""); const [max, setMax] = useState(rule.max ?? "");
  const asPct = fmt === FMT.PERCENT;
  const wrap = (v) => (v === "" ? undefined : asPct ? Number(v) / 100 : Number(v));
  const disp = (v) => (v == null || v === "" ? "" : asPct ? +(Number(v) * 100).toFixed(4) : v);
  const save = () => { const r = { dir, fmt }; if (dir === "range") { r.min = wrap(min); r.max = wrap(max); } else if (dir !== "info") { r.good = wrap(good); r.fair = wrap(fair); } onSave(r); };
  return (<Modal onClose={onClose} title={`${crit.name}: target`} desc={`${scopeName} · saved to this client only.`}>
    <div className="two"><div className="field"><label>Direction</label><select value={dir} onChange={(e) => setDir(e.target.value)}><option value="higher">Higher is better</option><option value="lower">Lower is better</option><option value="range">Within a band</option><option value="info">Track only</option></select></div>
      <div className="field"><label>Format</label><select value={fmt} onChange={(e) => setFmt(e.target.value)}><option value={FMT.PERCENT}>Percent</option><option value={FMT.DAYS}>Days</option><option value={FMT.COUNT}>Number</option></select></div></div>
    {dir === "range" ? <div className="two"><div className="field mono"><label>Band min {asPct ? "(%)" : ""}</label><input value={disp(min)} onChange={(e) => setMin(e.target.value)} /></div><div className="field mono"><label>Band max {asPct ? "(%)" : ""}</label><input value={disp(max)} onChange={(e) => setMax(e.target.value)} /></div></div>
      : dir !== "info" ? <div className="two"><div className="field mono"><label>On target {dir === "higher" ? "≥" : "≤"} {asPct ? "(%)" : ""}</label><input value={disp(good)} onChange={(e) => setGood(e.target.value)} /></div><div className="field mono"><label>At risk {dir === "higher" ? "≥" : "≤"} {asPct ? "(%)" : ""}</label><input value={disp(fair)} onChange={(e) => setFair(e.target.value)} /></div></div> : null}
    <Actions onClose={onClose} onOk={save} okLabel="Save target" />
  </Modal>);
}
function InfoModal({ data, onClose }) {
  const { crit, cell, rule } = data;
  const typed = cell?.method === "manual";
  const inputs = typed ? cell?.calc?.inputs : cell?.inputs;
  const val = typed && cell?.calc ? cell.calc.value : cell?.value;
  const isTxt = ["status", "yesno", "text"].includes(crit.type);
  return (<Modal onClose={onClose} title={crit.name} desc="How this number is calculated.">
    {crit.formula && <div className="field"><label>Formula</label><div style={{ fontSize: 13, fontFamily: "var(--font-mono)", background: "var(--line-2)", borderRadius: 8, padding: "9px 11px" }} className="mono">{crit.formula}</div></div>}
    {inputs && Object.keys(inputs).length > 0 && <div className="field"><label>This period's inputs</label>
      <div style={{ background: "var(--steel-soft)", borderRadius: 8, padding: "9px 11px", fontSize: 12.5 }}>
        {Object.entries(inputs).map(([k, v]) => <div key={k} style={{ display: "flex", justifyContent: "space-between", padding: "2px 0" }}><span style={{ color: "var(--ink-2)" }}>{k}</span><span className="mono" style={{ fontWeight: 600 }}>{typeof v === "number" ? v.toLocaleString() : v}</span></div>)}
        <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0 0", marginTop: 4, borderTop: "1px solid #cfe6df" }}><span style={{ fontWeight: 600, color: "var(--steel)" }}>Result</span><span className="mono" style={{ fontWeight: 600, color: "var(--steel)" }}>{isTxt ? val : fmtVal(val, crit.fmt || rule?.fmt)}</span></div>
      </div></div>}
    <div className="field"><label>Standard</label><div style={{ fontSize: 12.5, color: "var(--ink-2)" }}>{crit.std || "—"}</div></div>
    <div style={{ fontSize: 11.5, color: "var(--ink-3)" }}>{cell?.method === "upload" ? "Calculated automatically from the uploaded report." : cell?.method === "manual" ? (cell?.calc ? `A typed value is in use for this period: ${typeof cell.value === "number" ? fmtVal(cell.value, crit.fmt || rule?.fmt) : cell.value}. The figures above are the calculated ones.` : "Entered manually.") : ""}</div>
    <div className="modal-actions"><span /><button className="btn primary" onClick={onClose}>Got it</button></div>
  </Modal>);
}
// Popups are drawn straight onto the page body, so an animated container can never trap one mid-page.
function Overlay({ onClick, children }) {
  const node = <div className="overlay" onClick={onClick}>{children}</div>;
  return typeof document !== "undefined" && document.body ? createPortal(node, document.body) : node;
}
function Modal({ title, desc, children, onClose }) { return (<Overlay onClick={onClose}><div className="modal" onClick={(e) => e.stopPropagation()}><h3>{title}</h3>{desc && <p>{desc}</p>}{children}</div></Overlay>); }
function Actions({ onClose, onOk, okLabel, okDisabled }) { return (<div className="modal-actions"><button className="btn ghost" onClick={onClose}>Cancel</button><button className="btn primary" onClick={onOk} disabled={okDisabled}>{okLabel}</button></div>); }
