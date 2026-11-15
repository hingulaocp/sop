/* Lead-wise SOP web page: choose a day or all days, change what the joint check decides, download the SOP as Excel, save the changes.
   Figures come from compute.js (the same rules as the workbook); the Excel file from xlsx.js. Changes live in edits.json beside the page. */
(function () {
  "use strict";
  const S = window.SOPCompute, X = window.SOPXlsx;
  const app = document.getElementById("app");
  const esc0 = v => String(v === null || v === undefined ? "" : v).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  // ---------- theme: White / Black, remembered on this device
  function setTheme(t) { if (t) document.documentElement.setAttribute("data-theme", t); else document.documentElement.removeAttribute("data-theme"); try { if (t) localStorage.setItem("sop-theme", t); } catch (e) { /* none */ } }
  try { const t = localStorage.getItem("sop-theme"); if (t === "light" || t === "dark") setTheme(t); } catch (e) { /* none */ }
  const themeNow = () => document.documentElement.getAttribute("data-theme") || (matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
  const themeButtons = () => `<span class="themes" role="group" aria-label="Theme"><button data-theme-set="light" aria-pressed="${themeNow() === "light"}">White</button><button data-theme-set="dark" aria-pressed="${themeNow() === "dark"}">Black</button></span>`;
  const bindTheme = rerender => app.querySelectorAll("[data-theme-set]").forEach(b => b.addEventListener("click", () => { setTheme(b.dataset.themeSet); rerender(); }));

  // ---------- locked copy (GitHub Pages): the data travels encrypted (AES-GCM, key from the password by PBKDF2-SHA256); nothing shows without it
  async function unlock(pw) {
    const P = window.SOP_ENC; const b64 = x => Uint8Array.from(atob(x), c => c.charCodeAt(0));
    const base = await crypto.subtle.importKey("raw", new TextEncoder().encode(pw), "PBKDF2", false, ["deriveKey"]);
    const key = await crypto.subtle.deriveKey({ name: "PBKDF2", salt: b64(P.salt), iterations: P.iter, hash: "SHA-256" }, base, { name: "AES-GCM", length: 256 }, false, ["decrypt"]);
    const plain = await crypto.subtle.decrypt({ name: "AES-GCM", iv: b64(P.iv) }, key, b64(P.ct));
    const txt = await new Response(new Blob([plain]).stream().pipeThrough(new DecompressionStream("gzip"))).text();
    return JSON.parse(txt);
  }
  function lockScreen(msg) {
    app.innerHTML = `<header class="top"><div><h1>Lead-wise SOP</h1><div class="sub">Protected page. Enter the password to open it.</div></div><div class="actions">${themeButtons()}</div></header>
      <form class="card lock" id="lock" autocomplete="off"><label for="pw"><b>Password</b></label><input type="password" id="pw" autocomplete="current-password" required>
      <label class="switch" style="font-weight:500"><input type="checkbox" id="keep"> Keep it open in this tab</label>
      <button class="primary" type="submit">Open</button>${msg ? `<p class="bad-text" role="alert">${esc0(msg)}</p>` : ""}</form>`;
    bindTheme(() => lockScreen(msg));
    const f = document.getElementById("lock"); document.getElementById("pw").focus();
    f.addEventListener("submit", async ev => {
      ev.preventDefault(); const pw = document.getElementById("pw").value; const keep = document.getElementById("keep").checked;
      f.querySelector("button[type=submit]").disabled = true;
      try { const P = await unlock(pw); if (keep) { try { sessionStorage.setItem("sop-pw", pw); } catch (e) { /* none */ } } main(P.data, P.edits, true); }
      catch (e) { lockScreen("That password did not open the page."); }
    });
  }

  function main(DATA, PRE, VIEW) {
  const { SL, SH } = S;
  const HOSTED = !(window.claude && typeof window.claude.use === "function");   // true on a plain web host (GitHub Pages)
  const DAYS = DATA.days.slice().sort((a, b) => (a.day < b.day ? -1 : 1));
  const BYDAY = Object.fromEntries(DAYS.map(d => [d.day, d]));
  const esc = v => String(v === null || v === undefined ? "" : v).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const fmt = v => (v === null || v === undefined || v === "" ? "" : typeof v === "number" ? (Number.isInteger(v) ? v.toLocaleString("en-IN") : v.toFixed(3)) : esc(v));
  const pct = p => (p == null ? "—" : (100 * p).toFixed(1) + "%");
  const dmy = X.dmy, ist = X.ist;
  const short = d => dmy(d).slice(0, 5);
  const now = () => { const d = new Date(); const p = n => String(n).padStart(2, "0"); return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`; };

  // ---------- state
  let E = { v: 1, sign: { show: true, blocks: X.DEFAULT_SIGN.map(b => ({ ...b })) }, days: {}, log: [] };
  let saved = JSON.stringify(E);
  let tab = "sop", page = 0, filt = { q: "", shift: "", slab: "", edited: false };
  let caps = { artifact: null, downloads: null }, readOnly = false, busy = false, lastCopy = "";
  const dirty = () => JSON.stringify(E) !== saved;
  const sel = () => { const h = location.hash.replace("#", ""); const m = h.match(/^d-(\d{4}-\d{2}-\d{2})$/); return m && BYDAY[m[1]] ? [BYDAY[m[1]]] : DAYS; };
  const isAll = () => sel().length > 1;
  const dayE = d => E.days[d] || (E.days[d] = {});
  const results = () => sel().map(D => [D, S.computeDay(D, E.days[D.day])]);

  // set a value at a path inside E, keep a log entry that can be undone
  function getAt(obj, path) { return path.reduce((o, k) => (o === undefined || o === null ? undefined : o[k]), obj); }
  function setAt(obj, path, v) {
    let o = obj; for (let i = 0; i < path.length - 1; i++) { if (typeof o[path[i]] !== "object" || o[path[i]] === null) o[path[i]] = {}; o = o[path[i]]; }
    const k = path[path.length - 1]; if (v === undefined) delete o[k]; else o[k] = v;
  }
  function change(path, value, what, day) {
    const before = getAt(E, path); if (JSON.stringify(before) === JSON.stringify(value)) return;
    const last = E.log[E.log.length - 1];
    if (last && JSON.stringify(last.path) === JSON.stringify(path)) { last.after = value; last.what = what; last.t = now(); if (JSON.stringify(last.before) === JSON.stringify(value)) E.log.pop(); }
    else E.log.push({ t: now(), day: day || null, what, path, before, after: value });
    setAt(E, path, value); stash(); render();
  }
  function undo(i) { const l = E.log[i]; setAt(E, l.path, l.before); E.log.splice(i, 1); stash(); render(); }
  function stash() { try { localStorage.setItem("sop-draft", JSON.stringify({ base: saved, E })); } catch (e) { /* storage unavailable: changes stay in the page until saved */ } }
  function toast(t) { const d = document.createElement("div"); d.className = "toast"; d.textContent = t; d.setAttribute("role", "status"); document.body.appendChild(d); setTimeout(() => d.remove(), 3500); }

  // ---------- rendering
  function render() {
    const days = sel(); const R = results();
    const nChanges = E.log.length;
    let h = `<header class="top"><div><h1>NIT-185 lead-wise SOP · ${esc(DATA.site)}</h1>
      <div class="sub">${esc(DATA.contractor)} · OB tipper trips, GPS track against contractor · ${dmy(DAYS[0].day)} to ${dmy(DAYS[DAYS.length - 1].day)}</div></div>
      <div class="actions">${themeButtons()}${VIEW ? `<span class="pill">View only · changes are made by the owner</span><button class="primary" id="b-dl">Download Excel · ${days.length > 1 ? "all days" : short(days[0].day)}</button></div></header>` : `
        <span class="pill ${dirty() ? "edit" : "ok"}">${dirty() ? `${nChanges} change${nChanges === 1 ? "" : "s"} not saved` : nChanges ? `${nChanges} change${nChanges === 1 ? "" : "s"} saved` : "No changes"}</span>
        <button class="primary" id="b-dl">Download Excel · ${days.length > 1 ? "all days" : short(days[0].day)}</button>
        <button id="b-save" ${!dirty() || (readOnly && !HOSTED) || busy ? "disabled" : ""}>${HOSTED ? "Save changes (this browser + file)" : "Save changes"}</button>
        <button id="b-copy" ${nChanges ? "" : "disabled"}>Copy changes</button>
      </div></header>`}`;
    h += restoreBanner();
    h += `<nav class="days" aria-label="Report day"><a href="#all" aria-current="${isAll()}"><b>All days</b><small>${DAYS.length} days, in date order</small></a>`;
    for (const D of DAYS) {
      const r = S.computeDay(D, E.days[D.day]); const ed = E.log.some(l => l.day === D.day);
      h += `<a href="#d-${D.day}" aria-current="${!isAll() && days[0].day === D.day}"><b>${short(D.day)}${ed ? ' <span class="pill edit" title="changed on this page">edited</span>' : ""}</b><small>${fmt(r.Atot.final)} final · ${fmt(r.Atot.diff)} short</small></a>`;
    }
    h += `</nav>`;
    const tabs = [["sop", "SOP Daily Report"], ["just", "Justification Register"], ["trips", "Trips"], ["tw", "Tipper-wise"], ["rev", "For review"], ["nc", "Not counted"], ["sign", "Signatures"], ["chg", `Changes${nChanges ? " (" + nChanges + ")" : ""}`]];
    h += `<nav class="tabs" role="tablist">${tabs.map(([k, t]) => `<button role="tab" id="t-${k}" aria-selected="${tab === k}" data-tab="${k}">${t}</button>`).join("")}</nav>`;
    if (tab === "sop") h += (isAll() ? monthCard(R) : "") + R.map(([D, r]) => sopCard(D, r)).join("");
    else if (tab === "just") h += R.map(([D, r]) => justCard(D, r)).join("");
    else if (tab === "trips") h += tripsCard(R);
    else if (tab === "tw") h += twCard(R);
    else if (tab === "rev") h += revCard(R, "rev");
    else if (tab === "nc") h += revCard(R, "nc");
    else if (tab === "sign") h += signCard();
    else if (tab === "chg") h += changesCard();
    app.innerHTML = h; bind();
  }
  function restoreBanner() {
    if (VIEW) return "";
    let d = null; try { d = JSON.parse(localStorage.getItem("sop-draft") || "null"); } catch (e) { d = null; }
    if (!d || !d.E || d.base !== saved || JSON.stringify(d.E) === JSON.stringify(E) || dirty()) return "";
    return `<div class="banner"><span>You have ${d.E.log.length} change${d.E.log.length === 1 ? "" : "s"} from an earlier visit that were not saved.</span><button id="b-restore">Restore them</button><button id="b-forget">Discard them</button></div>`;
  }
  const head = cols => `<tr>${cols.map(c => `<th${c[1] ? ` class="${c[1]}"` : ""}>${esc(c[0] || c)}</th>`).join("")}</tr>`;
  const H = a => head(a.map(x => (Array.isArray(x) ? x : [x])));

  function monthCard(R) {
    let rows = "", tot = Array(6 + SL.length).fill(0), dK = 0, uK = 0, nK = 0;
    for (const [D, r] of R) {
      const v = [r.Atot.reported, r.Atot.vts, r.Atot.gps, r.Atot.justified, r.Atot.final, r.Atot.diff, ...r.Btot.slabs]; v.forEach((x, i) => (tot[i] += x));
      const u = r.Ctot.und; if (u != null) { dK += r.Ctot.total; uK += u; nK++; }
      const cells = v.map(fmt); cells.splice(2, 0, u == null ? "—" : fmt(u), pct(r.Ctot.pctUnd));
      rows += `<tr><td class="l"><a href="#d-${D.day}">${dmy(D.day)}</a></td>${cells.map(x => `<td>${x}</td>`).join("")}</tr>`;
    }
    const tc = tot.map(fmt); tc.splice(2, 0, nK ? fmt(uK) : "—", nK ? pct(dK + uK ? uK / (dK + uK) : 0) : "—");
    rows += `<tr class="tot"><td class="l">TOTAL</td>${tc.map(x => `<td>${x}</td>`).join("")}</tr>`;
    return `<section class="card"><h2>Month so far</h2><div class="scroll"><table>${H([["Date", "l"], "Reported by contractor", "VTS designated", "VTS undesignated", "% undesignated", "GPS hauls accepted", "Trips given — justified", "FINAL TRIPS", "Difference (reported − final)", ...SL.map(s => "Final " + s)])}${rows}</table></div>
      <p class="note">VTS designated / undesignated: NIT-185's own trips (Transporter AJANTHA-BLAPPL (JV)) in the VTS Trip List; % undesignated = undesignated ÷ (designated + undesignated), as in HQ's NIT Lead Report. — : that day's Trip List export had designated trips only; the TOTAL % is over the days where it is known.</p>
      <p class="note">Each day below in date order. Pick a day above to work on it alone; the Excel download follows your choice (one day, or all days with a Month sheet).</p></section>`;
  }
  function sopCard(D, r) {
    const de = E.days[D.day] || {};
    const a = r.A.map(x => `<tr><td class="l">${SH[x.shift]}</td><td>${fmt(x.reported)}</td><td>${fmt(x.vts)}</td><td>${fmt(x.gps)}</td><td>${fmt(x.review)}</td><td>${fmt(x.noGpsRep)}</td><td>${fmt(x.justified)}</td><td><b>${fmt(x.final)}</b></td><td class="${x.diff < 0 ? "miss" : ""}">${fmt(x.diff)}</td></tr>`).join("");
    const t = r.Atot;
    const b = r.B.map(x => `<tr><td class="l">${SH[x.shift]}</td>${x.slabs.map(v => `<td>${fmt(v)}</td>`).join("")}<td>${fmt(x.total)}</td><td>${x.ok ? '<span class="pill ok">OK</span>' : '<span class="pill bad">MISMATCH</span>'}</td></tr>`).join("");
    const c = r.C.map(x => `<tr><td class="l">${SH[x.shift]}</td>${x.slabs.map(v => `<td>${fmt(v)}</td>`).join("")}<td>${fmt(x.total)}</td><td>${x.und == null ? "—" : fmt(x.und)}</td><td>${pct(x.pctUnd)}</td></tr>`).join("");
    const d = r.D.map(x => `<tr><td class="l">${SH[x.shift]}</td>${[...x.con, ...x.fin].map(v => `<td>${fmt(v)}</td>`).join("")}${x.diff.map(v => `<td>${fmt(v)}</td>`).join("")}</tr>`).join("");
    return `<section class="card daycard"><div class="head"><h2>${dmy(D.day)} · SOP Daily Report</h2>
      <span><label class="switch"><input type="checkbox" data-act="justified" data-day="${D.day}" ${r.justifiedOn ? "checked" : ""}> Include trips given — justified</label>
      <label class="switch">Cap <select data-act="capby" data-day="${D.day}" aria-label="Cap on justified trips">${[["shift", "Per shift"], ["day", "Per day"], ["none", "No cap"]].map(([v, t]) => `<option value="${v}" ${r.capBy === v ? "selected" : ""}>${t}</option>`).join("")}</select></label></span></div>
      <p class="note">Report day 05:00 → next day 05:00. Final trips = GPS hauls accepted + trips given for tipper-shifts with no GPS data, ${r.capBy === "none" ? "given as reported (no cap)" : "capped at reported − GPS hauls " + (r.capBy === "shift" ? `of each shift (A ${fmt(r.capShift.A)} · B ${fmt(r.capShift.B)} · C ${fmt(r.capShift.C)}), and never above the day (${fmt(r.cap)})` : `of the day (${fmt(r.cap)}), filled A → B → C`)}.</p>
      <h3>A. Trips: contractor, VTS and GPS</h3>
      <div class="scroll"><table>${H([["Shift", "l"], "Reported by contractor", "VTS designated", "GPS hauls accepted", "Under review (not counted)", "No-GPS trips reported", "Trips given — justified", "FINAL TRIPS", "Difference (reported − final)"])}${a}
      <tr class="tot"><td class="l">DAY TOTAL</td><td>${fmt(t.reported)}</td><td>${fmt(t.vts)}</td><td>${fmt(t.gps)}</td><td>${fmt(t.review)}</td><td>${fmt(t.noGpsRep)}</td><td>${fmt(t.justified)}</td><td>${fmt(t.final)}</td><td>${fmt(t.diff)}</td></tr></table></div>
      <h3>B. Final trips by lead slab (km), lower &lt; lead ≤ upper</h3>
      <div class="scroll"><table>${H([["Shift", "l"], ...SL, "Total", "Total = Final?"])}${b}<tr class="tot"><td class="l">DAY TOTAL</td>${r.Btot.slabs.map(v => `<td>${fmt(v)}</td>`).join("")}<td>${fmt(r.Btot.total)}</td><td>${r.Btot.ok ? '<span class="pill ok">OK</span>' : '<span class="pill bad">MISMATCH</span>'}</td></tr></table></div>
      <div class="grid2"><div><h3>C. VTS designated trips by VTS distance, and undesignated</h3><div class="scroll"><table>${H([["Shift", "l"], ...SL, "Total", "VTS undesignated", "% undesignated"])}${c}<tr class="tot"><td class="l">DAY TOTAL</td>${r.Ctot.slabs.map(v => `<td>${fmt(v)}</td>`).join("")}<td>${fmt(r.Ctot.total)}</td><td>${r.Ctot.und == null ? "—" : fmt(r.Ctot.und)}</td><td>${pct(r.Ctot.pctUnd)}</td></tr></table></div></div></div>
      <h3>D. Lead-wise: contractor's reported trips against final trips</h3>
      <div class="scroll"><table>${H([["Shift", "l"], ...SL.map(s => s + " contractor"), ...SL.map(s => s + " final"), ...SL.map(s => s + " final − contr.")])}${d}
      <tr class="tot"><td class="l">DAY TOTAL</td>${[...r.Dtot.con, ...r.Dtot.fin, ...r.Dtot.diff].map(v => `<td>${fmt(v)}</td>`).join("")}</tr></table></div>
      ${E.sign.show ? sigBlock() : '<p class="note">Signature blocks are switched off (tab Signatures); the download has none.</p>'}</section>`;
  }
  function sigBlock() { return `<div class="sigs">${E.sign.blocks.map(b => `<div class="sig"><div class="line">Signature · Name · Date</div><div class="role">${esc(b.role)}</div><div class="org">${esc(b.org)}</div></div>`).join("")}</div>`; }
  function justCard(D, r) {
    const rows = r.register.map(j => {
      const ed = (E.days[D.day] || {}).jr && E.days[D.day].jr[j.key];
      return `<tr class="${ed ? "edited" : ""}"><td>${j.sl}</td><td class="l">${j.shift}</td><td class="l">${esc(j.fleet)}</td><td class="l">${esc(j.reg || "")}</td><td>${fmt(j.reported)}</td><td>${fmt(j.gps)}</td><td>${fmt(j.diff)}</td><td><b>${fmt(j.included)}</b></td>
        <td><select class="slab" aria-label="Slab for ${esc(j.fleet)} shift ${j.shift}" data-act="jr-slab" data-day="${D.day}" data-key="${esc(j.key)}">${SL.map(s => `<option ${s === j.slab ? "selected" : ""}>${s}</option>`).join("")}</select></td>
        <td class="l note">${esc(j.reason)}</td><td><input type="text" aria-label="Justification" data-act="jr-just" data-day="${D.day}" data-key="${esc(j.key)}" value="${esc(j.just)}" placeholder="Write the justification"></td>
        <td><input type="text" aria-label="Bound register page" style="min-width:70px" data-act="jr-page" data-day="${D.day}" data-key="${esc(j.key)}" value="${esc(j.page)}" placeholder="Page"></td></tr>`;
    }).join("");
    const s = k => r.register.reduce((a, j) => a + j[k], 0);
    return `<section class="card daycard"><div class="head"><h2>${dmy(D.day)} · Justification Register</h2><span class="note">${r.capBy === "none" ? "No cap: the trips reported are given as they are" : r.capBy === "shift" ? `Cap per shift: A ${fmt(r.capShift.A)} · B ${fmt(r.capShift.B)} · C ${fmt(r.capShift.C)} (reported − GPS hauls of the shift; day ${fmt(r.cap)} at most)` : `Cap ${fmt(r.cap)} for the day (reported − GPS hauls), filled A → B → C`}${r.justifiedOn ? "" : " · switched OFF for this day"}</span></div>
      ${r.register.length ? `<div class="scroll"><table>${H(["Sl.", ["Shift", "l"], ["Fleet No", "l"], ["Reg No", "l"], "Reported", "GPS hauls", "Difference", "Trips included", "Slab", ["Reason", "l"], ["Justification", "l"], "Register page"])}${rows}
      <tr class="tot"><td class="l" colspan="4">TOTAL</td><td>${fmt(s("reported"))}</td><td>${fmt(s("gps"))}</td><td>${fmt(s("diff"))}</td><td>${fmt(s("included"))}</td><td colspan="4"></td></tr></table></div>` : `<p class="muted">No trip was given on justification on this day.</p>`}
      <p class="note">Slab defaults to 0–1. Choose another and the SOP Daily Report follows. Only technical mismatches may be included (SOP 7).</p>
      ${E.sign.show ? sigBlock() : ""}</section>`;
  }
  function allCounted(R) { return R.flatMap(([D, r]) => r.counted.map(t => ({ ...t, day: D.day }))); }
  function tripsCard(R) {
    let rows = allCounted(R);
    const q = filt.q.trim().toLowerCase();
    if (q) rows = rows.filter(t => String(t.No) === q || t.Fleet.toLowerCase().includes(q) || (t.from + " " + t.to).toLowerCase().includes(q));
    if (filt.shift) rows = rows.filter(t => t.Shift === filt.shift);
    if (filt.slab) rows = rows.filter(t => t.slab === filt.slab);
    if (filt.edited) rows = rows.filter(t => t.edited);
    const n = rows.length, per = 200, pages = Math.max(1, Math.ceil(n / per)); page = Math.min(page, pages - 1);
    const body = rows.slice(page * per, page * per + per).map(t => {
      const de = (E.days[t.day] || {}).trips || {}; const e = de[t.No];
      return `<tr class="${e ? "edited" : ""}">${isAll() ? `<td class="l">${short(t.day)}</td>` : ""}<td>${t.No}</td><td class="l">${t.Shift}</td><td class="l">${esc(t.Fleet)}</td><td class="l">${ist(t.ld).slice(11)}</td><td class="l">${ist(t.tp).slice(11)}</td>
        <td class="l">${esc(t.from)}</td><td class="l">${esc(t.to)}</td><td>${t.lead === null ? "" : t.lead.toFixed(3)}</td>
        <td><select class="slab" aria-label="Slab of trip ${t.No}" data-act="trip-slab" data-day="${t.day}" data-no="${t.No}">${SL.map(s => `<option ${s === t.slab ? "selected" : ""}>${s}</option>`).join("")}</select></td>
        <td class="l note">${esc(t.note)}</td><td><button data-act="trip-out" data-day="${t.day}" data-no="${t.No}" title="Leave this trip out of the count">Don't count</button></td></tr>`;
    }).join("");
    return `<section class="card"><div class="row"><h2>Trips counted${isAll() ? ", all days" : " · " + dmy(sel()[0].day)}</h2><span class="muted">${fmt(n)} trip${n === 1 ? "" : "s"}</span></div>
      <div class="row"><input type="text" id="f-q" style="max-width:260px" placeholder="Trip No, tipper or place" value="${esc(filt.q)}" aria-label="Search trips">
        <select id="f-shift" aria-label="Shift"><option value="">All shifts</option>${["A", "B", "C"].map(s => `<option ${filt.shift === s ? "selected" : ""}>${s}</option>`).join("")}</select>
        <select id="f-slab" aria-label="Slab"><option value="">All slabs</option>${SL.map(s => `<option ${filt.slab === s ? "selected" : ""}>${s}</option>`).join("")}</select>
        <label class="switch" style="font-weight:500"><input type="checkbox" id="f-ed" ${filt.edited ? "checked" : ""}> Changed only</label></div>
      <p class="note">Times are where the lead was measured (IST). Change a slab or leave a trip out; trips left out move to “Not counted”, where they can be counted again.</p>
      <div class="scroll"><table>${H([...(isAll() ? [["Date", "l"]] : []), "No", ["Shift", "l"], ["Tipper", "l"], ["Leaves", "l"], ["Tips", "l"], ["From", "l"], ["To", "l"], "Lead (km)", "Slab", ["Note", "l"], ""])}${body}</table></div>
      <div class="pager"><button data-act="pg" data-p="${page - 1}" ${page <= 0 ? "disabled" : ""}>Previous</button><span>Page ${page + 1} of ${pages}</span><button data-act="pg" data-p="${page + 1}" ${page >= pages - 1 ? "disabled" : ""}>Next</button></div></section>`;
  }
  function twCard(R) {
    const rows = R.flatMap(([D, r]) => r.tw.map(t => ({ ...t, day: D.day })));
    const body = rows.map(t => `<tr>${isAll() ? `<td class="l">${short(t.day)}</td>` : ""}<td class="l">${t.shift}</td><td class="l">${esc(t.fleet)}</td><td>${fmt(t.rep)}</td><td>${t.repLead === null ? "" : t.repLead.toFixed(2)}</td><td>${esc(t.repSlab || "")}</td><td>${fmt(t.gps)}</td>
      <td class="l">${esc(t.bySlab)}</td><td>${t.gpsLead === null ? "" : t.gpsLead.toFixed(2)}</td><td class="${t.diff < 0 ? "miss" : ""}">${fmt(t.diff)}</td><td class="l">${esc(t.chk)}</td><td class="l note">${esc(t.rem)}</td></tr>`).join("");
    return `<section class="card"><h2>Tipper-wise${isAll() ? ", all days" : " · " + dmy(sel()[0].day)}</h2>
      <div class="scroll"><table>${H([...(isAll() ? [["Date", "l"]] : []), ["Shift", "l"], ["Tipper", "l"], "Reported", "Reported lead", "Reported slab", "GPS trips", ["GPS by slab", "l"], "GPS lead (median)", "Reported − GPS", ["Slab check", "l"], ["Remark", "l"]])}${body}
      <tr class="tot"><td class="l" colspan="${isAll() ? 3 : 2}">TOTAL (${fmt(rows.length)} tipper-shifts)</td><td>${fmt(rows.reduce((a, t) => a + t.rep, 0))}</td><td></td><td></td><td>${fmt(rows.reduce((a, t) => a + t.gps, 0))}</td><td></td><td></td><td>${fmt(rows.reduce((a, t) => a + t.diff, 0))}</td><td colspan="2" class="l note">Reported − GPS is 0 for no-GPS tipper-shifts (their trips are in the Justification Register)</td></tr></table></div></section>`;
  }
  function revCard(R, which) {
    const N = S.N;
    let body = "";
    for (const [D, r] of R) {
      const te = (E.days[D.day] || {}).trips || {};
      if (which === "rev") body += r.forReview.map(t => `<tr>${isAll() ? `<td class="l">${short(D.day)}</td>` : ""}<td>${t.No}</td><td class="l">${t.Shift}</td><td class="l">${esc(t.Fleet)}</td><td class="l">${ist(t.ld).slice(11)}</td><td class="l">${ist(t.tp).slice(11)}</td><td class="l">${esc(t.from)}</td><td class="l">${esc(t.to)}</td><td>${t.lead === null || t.lead === undefined ? "" : t.lead.toFixed(3)}</td><td class="l">${esc(t.why)}</td><td class="l">${esc(t.countedTxt)}</td>
        <td>${t.countedTxt === "No" ? `<button data-act="trip-in" data-day="${D.day}" data-no="${t.No}">Count it</button>` : `<button data-act="trip-out" data-day="${D.day}" data-no="${t.No}">Don't count</button>`}</td></tr>`).join("");
      else {
        const out = D.trips.filter(t => te[t[0]] && te[t[0]].count === false).map(t => [t[0], "Left out on this page", t[1], t[2], t[3], t[4], t[5], t[6], t[7], "gps"]);
        body += out.concat(r.notCounted.map(x => [...x, "nc"])).map(x => `<tr class="${x[9] === "gps" ? "edited" : ""}">${isAll() ? `<td class="l">${short(D.day)}</td>` : ""}<td>${x[N.No]}</td><td class="l">${esc(x[N.why])}</td><td class="l">${x[N.Shift]}</td><td class="l">${esc(x[N.Fleet])}</td><td class="l">${ist(x[N.ld]).slice(11)}</td><td class="l">${ist(x[N.tp]).slice(11)}</td><td class="l">${esc(x[N.from])}</td><td class="l">${esc(x[N.to])}</td><td>${x[N.lead] === null ? "" : Number(x[N.lead]).toFixed(3)}</td>
          <td>${x[9] === "gps" ? `<button data-act="trip-undo" data-day="${D.day}" data-no="${x[N.No]}">Count again</button>` : `<button data-act="trip-in" data-day="${D.day}" data-no="${x[N.No]}">Count it</button>`}</td></tr>`).join("");
      }
    }
    const hd = which === "rev" ? [...(isAll() ? [["Date", "l"]] : []), "No", ["Shift", "l"], ["Tipper", "l"], ["Leaves", "l"], ["Tips", "l"], ["From", "l"], ["To", "l"], "Lead (km)", ["Why", "l"], ["Counted?", "l"], ""]
      : [...(isAll() ? [["Date", "l"]] : []), "No", ["Why", "l"], ["Shift", "l"], ["Tipper", "l"], ["Leaves", "l"], ["Arrives", "l"], ["From", "l"], ["To", "l"], "Odometer lead (km)", ""];
    return `<section class="card"><h2>${which === "rev" ? "For review" : "Not counted"}${isAll() ? ", all days" : " · " + dmy(sel()[0].day)}</h2>
      <p class="note">${which === "rev" ? "Leads over 5 km are counted in 4–5 and listed here, with any haul still under review." : "Runs that are not trips under the agreed rules (parking runs, runs starting in a dump, a trip's part after 05:00), and any trip left out on this page. “Count it” adds the run in the slab of its odometer lead; change the slab under Trips."}</p>
      <div class="scroll"><table>${H(hd)}${body || `<tr><td class="l" colspan="12">Nothing here.</td></tr>`}</table></div></section>`;
  }
  function signCard() {
    const bl = E.sign.blocks.map((b, i) => `<div class="sig"><label class="muted" for="sr-${i}">Role</label><input type="text" id="sr-${i}" data-act="sig-role" data-i="${i}" value="${esc(b.role)}">
      <label class="muted" for="so-${i}">Place / firm</label><input type="text" id="so-${i}" data-act="sig-org" data-i="${i}" value="${esc(b.org)}"><button data-act="sig-del" data-i="${i}">Remove this block</button></div>`).join("");
    return `<section class="card"><h2>Signatures</h2>
      <label class="switch"><input type="checkbox" data-act="sig-show" ${E.sign.show ? "checked" : ""}> Print signature blocks on the SOP Daily Report and the Justification Register</label>
      <p class="note">These blocks are used on every day and in the download. Names and dates are left blank, to be written by hand.</p>
      <div class="sigs">${bl}</div><div class="row"><button data-act="sig-add">Add a block</button><button data-act="sig-reset">Back to the SOP's four</button></div></section>`;
  }
  function changesCard() {
    const rows = E.log.map((l, i) => `<tr><td class="l">${esc(l.t)}</td><td class="l">${l.day ? dmy(l.day) : "all days"}</td><td class="l">${esc(l.what)}</td><td><button data-act="undo" data-i="${i}">Undo</button></td></tr>`).reverse().join("");
    return `<section class="card"><h2>Changes</h2>
      <p class="note">${VIEW ? "Changes made by the owner on the original page; they are already in every figure here and in the download." : "“Save changes” keeps them with this page for the next visit."}${VIEW ? "" : ` “Copy changes” copies them as text to hand to Claude, so the same changes go into the reconciliation records in git. The download lists them in its Changes sheet.`}</p>
      <div class="scroll"><table>${H([["When", "l"], ["Date", "l"], ["Change", "l"], ""])}${rows || `<tr><td class="l" colspan="4">No changes yet.</td></tr>`}</table></div>
      ${E.log.length ? '<div class="row"><button data-act="undo-all">Undo all changes not saved</button></div>' : ""}
      ${lastCopy ? `<label class="muted" for="copybox">If copying did not work, select this text and copy it:</label><textarea class="copy" id="copybox" readonly>${esc(lastCopy)}</textarea>` : ""}</section>`;
  }

  // ---------- events
  function findTrip(day, no) {
    const D = BYDAY[day]; no = Number(no);
    const t = D.trips.find(x => x[0] === no); if (t) return { lead: t[7], slab: t[16], fleet: t[2], shift: t[1], src: "gps" };
    const r = D.review.find(x => x[0] === no) || D.notCounted.find(x => x[0] === no);
    if (!r) return null; const isRev = D.review.includes(r);
    return { lead: isRev ? r[7] : r[8], slab: S.slab(isRev ? r[7] : r[8]) || "0-1", fleet: isRev ? r[2] : r[3], shift: isRev ? r[1] : r[2], src: isRev ? "review" : "nc" };
  }
  function bind() {
    bindTheme(render);
    if (VIEW) {
      app.querySelectorAll("[data-tab]").forEach(b => b.addEventListener("click", () => { tab = b.dataset.tab; page = 0; render(); }));
      app.querySelectorAll('[data-act]').forEach(el => { if (el.dataset.act === "pg") el.addEventListener("click", () => { page = Number(el.dataset.p); render(); }); else if (el.tagName === "BUTTON") el.remove(); else el.disabled = true; });
      ["f-q", "f-shift", "f-slab", "f-ed"].forEach(id => { const el = document.getElementById(id); if (el) el.addEventListener("change", () => { filt.q = (document.getElementById("f-q") || {}).value || ""; filt.shift = (document.getElementById("f-shift") || {}).value || ""; filt.slab = (document.getElementById("f-slab") || {}).value || ""; filt.edited = !!(document.getElementById("f-ed") || {}).checked; page = 0; render(); }); });
      document.getElementById("b-dl").addEventListener("click", download); return;
    }
    app.querySelectorAll("[data-tab]").forEach(b => b.addEventListener("click", () => { tab = b.dataset.tab; page = 0; render(); }));
    const on = (sel, ev, fn) => app.querySelectorAll(sel).forEach(el => el.addEventListener(ev, () => fn(el)));
    on('[data-act="justified"]', "change", el => change(["days", el.dataset.day, "justified"], el.checked ? undefined : false, `${short(el.dataset.day)}: trips given — justified ${el.checked ? "included" : "switched off"}`, el.dataset.day));
    on('[data-act="capby"]', "change", el => change(["days", el.dataset.day, "cap"], el.value === "shift" ? undefined : el.value, `${short(el.dataset.day)}: justified trips ${el.value === "none" ? "as reported (no cap)" : "capped per " + el.value}`, el.dataset.day));
    on('[data-act="jr-slab"]', "change", el => change(["days", el.dataset.day, "jr", el.dataset.key, "slab"], el.value === "0-1" ? undefined : el.value, `${el.dataset.key.split("|")[1]} shift ${el.dataset.key.split("|")[0]}: justified trips in slab ${el.value}`, el.dataset.day));
    on('[data-act="jr-just"]', "change", el => change(["days", el.dataset.day, "jr", el.dataset.key, "just"], el.value.trim() || undefined, `${el.dataset.key.split("|")[1]} shift ${el.dataset.key.split("|")[0]}: justification “${el.value.trim()}”`, el.dataset.day));
    on('[data-act="jr-page"]', "change", el => change(["days", el.dataset.day, "jr", el.dataset.key, "page"], el.value.trim() || undefined, `${el.dataset.key.split("|")[1]} shift ${el.dataset.key.split("|")[0]}: bound register page ${el.value.trim()}`, el.dataset.day));
    on('[data-act="trip-slab"]', "change", el => {
      const t = findTrip(el.dataset.day, el.dataset.no); const cur = ((E.days[el.dataset.day] || {}).trips || {})[el.dataset.no] || {};
      const v = { ...cur, slab: el.value === t.slab && t.src === "gps" ? undefined : el.value }; if (v.slab === undefined) delete v.slab;
      change(["days", el.dataset.day, "trips", el.dataset.no], Object.keys(v).length ? v : undefined, `Trip ${el.dataset.no} (${t.fleet}, ${t.shift}): slab ${el.value}`, el.dataset.day);
    });
    on('[data-act="trip-out"]', "click", el => { const t = findTrip(el.dataset.day, el.dataset.no); change(["days", el.dataset.day, "trips", el.dataset.no], t.src === "gps" ? { count: false } : undefined, `Trip ${el.dataset.no} (${t.fleet}, ${t.shift}): not counted`, el.dataset.day); });
    on('[data-act="trip-in"]', "click", el => { const t = findTrip(el.dataset.day, el.dataset.no); change(["days", el.dataset.day, "trips", el.dataset.no], { count: true, slab: t.slab }, `Trip ${el.dataset.no} (${t.fleet}, ${t.shift}): counted, slab ${t.slab}`, el.dataset.day); });
    on('[data-act="trip-undo"]', "click", el => { const t = findTrip(el.dataset.day, el.dataset.no); change(["days", el.dataset.day, "trips", el.dataset.no], undefined, `Trip ${el.dataset.no} (${t.fleet}, ${t.shift}): counted again`, el.dataset.day); });
    on('[data-act="pg"]', "click", el => { page = Number(el.dataset.p); render(); });
    on('[data-act="sig-show"]', "change", el => change(["sign", "show"], el.checked, `Signature blocks ${el.checked ? "printed" : "removed"}`));
    on('[data-act="sig-role"]', "change", el => { const b = E.sign.blocks.map(x => ({ ...x })); b[+el.dataset.i].role = el.value; change(["sign", "blocks"], b, `Signature block ${+el.dataset.i + 1}: role “${el.value}”`); });
    on('[data-act="sig-org"]', "change", el => { const b = E.sign.blocks.map(x => ({ ...x })); b[+el.dataset.i].org = el.value; change(["sign", "blocks"], b, `Signature block ${+el.dataset.i + 1}: “${el.value}”`); });
    on('[data-act="sig-del"]', "click", el => { const b = E.sign.blocks.filter((_, i) => i !== +el.dataset.i); change(["sign", "blocks"], b, `Signature block “${E.sign.blocks[+el.dataset.i].role}” removed`); });
    on('[data-act="sig-add"]', "click", () => change(["sign", "blocks"], [...E.sign.blocks.map(x => ({ ...x })), { role: "Designation", org: "Hingula OCP" }], "Signature block added"));
    on('[data-act="sig-reset"]', "click", () => change(["sign", "blocks"], X.DEFAULT_SIGN.map(b => ({ ...b })), "Signature blocks back to the SOP's four"));
    on('[data-act="undo"]', "click", el => undo(+el.dataset.i));
    on('[data-act="undo-all"]', "click", () => { E = JSON.parse(saved); stash(); render(); });
    const fq = document.getElementById("f-q"); if (fq) fq.addEventListener("change", () => { filt.q = fq.value; page = 0; render(); });
    const fs = document.getElementById("f-shift"); if (fs) fs.addEventListener("change", () => { filt.shift = fs.value; page = 0; render(); });
    const fb = document.getElementById("f-slab"); if (fb) fb.addEventListener("change", () => { filt.slab = fb.value; page = 0; render(); });
    const fe = document.getElementById("f-ed"); if (fe) fe.addEventListener("change", () => { filt.edited = fe.checked; page = 0; render(); });
    const br = document.getElementById("b-restore"); if (br) br.addEventListener("click", () => { try { E = JSON.parse(localStorage.getItem("sop-draft")).E; } catch (e) { /* keep current */ } render(); });
    const bf = document.getElementById("b-forget"); if (bf) bf.addEventListener("click", () => { try { localStorage.removeItem("sop-draft"); } catch (e) { /* none */ } render(); });
    document.getElementById("b-dl").addEventListener("click", download);
    document.getElementById("b-save").addEventListener("click", save);
    document.getElementById("b-copy").addEventListener("click", copy);
  }
  async function download() {
    const days = sel();
    if (!window.ExcelJS) { toast("The Excel library did not load. Check the connection and reload the page."); return; }
    const plain = !caps.downloads;   // outside claude.ai (e.g. GitHub Pages): an ordinary browser download
    try {
      const wb = X.buildWorkbook(window.ExcelJS, DATA, days, E);
      const buf = await wb.xlsx.writeBuffer();
      const name = days.length > 1 ? `SOP_Leadwise_NIT-185_${short(days[0].day)}-${dmy(days[days.length - 1].day)}.xlsx` : `SOP_Leadwise_NIT-185_${dmy(days[0].day)}.xlsx`;
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      if (plain) { const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000); }
      else await caps.downloads.save({ filename: name, data: blob });
      toast("Saved " + name);
    } catch (e) {
      if (e && e.code === "declined") return;
      toast(e && e.code === "rate_limited" ? "A download is already waiting for your answer." : "The download did not go through" + (e && e.code ? ` (${e.code})` : "") + ".");
    }
  }
  async function save() {
    if (!caps.artifact) {
      if (HOSTED) {   // plain web host: keep the changes in this browser and hand them over as a file for Claude / git
        try { localStorage.setItem("sop-saved", JSON.stringify(E)); } catch (e) { /* none */ }
        saved = JSON.stringify(E); try { localStorage.removeItem("sop-draft"); } catch (e) { /* none */ }
        const blob = new Blob([JSON.stringify(E, null, 1)], { type: "application/json" }); const a = document.createElement("a");
        a.href = URL.createObjectURL(blob); a.download = "sop-changes-" + now().replace(/[.: ]/g, "") + ".json"; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
        toast("Kept in this browser, and saved as a file: send it to Claude to put the changes in git."); render(); return;
      }
      toast("Saving is not available in this view. Use “Copy changes” and give the text to Claude."); return;
    }
    busy = true; render();
    const text = JSON.stringify(E, null, 1);
    try { sessionStorage.setItem("sop-draft-session", text); } catch (e) { /* none */ }
    try {
      await caps.artifact.publish({ "edits.json": { content: text, contentType: "application/json" } });
      saved = JSON.stringify(E); try { localStorage.removeItem("sop-draft"); } catch (e) { /* none */ }
      toast("Changes saved with the page.");
    } catch (e) {
      const c = e && e.code;
      if (c === "not_writer" || c === "not_granted" || c === "capability_disabled" || c === "consent_required" || c === "not_declared") { readOnly = true; toast("This view cannot save. Use “Copy changes” and give the text to Claude."); }
      else if (c === "conflict") toast("The page was changed elsewhere; it reloads with that version. Restore your changes from the banner.");
      else toast("Saving did not go through" + (c ? ` (${c})` : "") + ". Your changes are still here.");
    }
    busy = false; render();
  }
  async function copy() {
    lastCopy = JSON.stringify({ page: "NIT-185 lead-wise SOP", changes: E.log.map(l => ({ when: l.t, day: l.day, change: l.what, path: l.path, value: l.after })), edits: E });
    try { await navigator.clipboard.writeText(lastCopy); toast("Changes copied. Paste them to Claude."); }
    catch (e) { tab = "chg"; toast("Copying was refused here: select the text under Changes and copy it."); }
    render();
  }

  // ---------- start: saved changes, then capabilities
  async function start() {
    if (VIEW) {
      if (PRE && PRE.v === 1) E = { ...E, ...PRE, sign: { ...E.sign, ...(PRE.sign || {}) } };
      saved = JSON.stringify(E); readOnly = true;
      window.addEventListener("hashchange", () => { page = 0; render(); }); render(); return;
    }
    try {
      const r = await fetch("edits.json", { cache: "no-store" });
      if (r.ok) { const j = await r.json(); if (j && j.v === 1) { E = { ...E, ...j, sign: { ...E.sign, ...(j.sign || {}) } }; } }
    } catch (e) { /* no saved changes yet */ }
    if (HOSTED) { try { const k = JSON.parse(localStorage.getItem("sop-saved") || "null"); if (k && k.v === 1) E = k; } catch (e) { /* none */ } }
    saved = JSON.stringify(E);
    try { const s = sessionStorage.getItem("sop-draft-session"); if (s && s === JSON.stringify(JSON.parse(s)) && JSON.stringify(JSON.parse(s)) === saved) sessionStorage.removeItem("sop-draft-session"); } catch (e) { /* none */ }
    window.addEventListener("hashchange", () => { page = 0; render(); });
    render();
    if (window.claude && typeof window.claude.use === "function") {
      window.claude.use("downloads").then(d => { caps.downloads = d; }).catch(() => {});
      window.claude.use("artifact").then(a => { caps.artifact = a; if (!a) readOnly = true; render(); }).catch(() => { readOnly = true; render(); });
    }
  }
  start();
  }   // main

  if (window.SOP_DATA) main(window.SOP_DATA, null, false);          // claude.ai: the owner's page, changes saved with it
  else if (window.SOP_ENC) {                                          // GitHub Pages: locked, view and download only
    let pw = null; try { pw = sessionStorage.getItem("sop-pw"); } catch (e) { /* none */ }
    if (pw) unlock(pw).then(P => main(P.data, P.edits, true)).catch(() => lockScreen()); else lockScreen();
  } else app.innerHTML = '<p class="muted">The data did not load. Reload the page.</p>';
})();
