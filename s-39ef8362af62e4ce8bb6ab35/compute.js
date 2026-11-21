/* Lead-wise SOP: every figure of the day report, rebuilt from the day's rows and the viewer's edits.
   Same rules as lab/leadwise-pipeline/build_report.py (the workbook): run test_compute.js after any change here.
   Works in the page (window.SOPCompute) and in node (require). */
(function (root) {
  "use strict";
  const SL = ["0-1", "1-2", "2-3", "3-4", "4-5"];
  const SH = { A: "A (05:00–13:00)", B: "B (13:00–21:00)", C: "C (21:00–05:00)" };
  const SHIFTS = ["A", "B", "C"];
  // lower < lead <= upper; a lead over 5 km is counted in 4–5 (the NIT's highest slab)
  function slab(x) {
    if (x === null || x === undefined || Number.isNaN(x)) return null;
    return x > 0 ? SL[Math.min(4, Math.max(0, Math.ceil(x) - 1))] : SL[0];
  }
  const zero = () => SL.map(() => 0);
  const T = { No: 0, Shift: 1, Fleet: 2, ld: 3, tp: 4, from: 5, to: 6, lead: 7, oneWay: 8, portal: 9, a: 10, b: 11, M: 12, c: 13, d: 14, note: 15, slab: 16 };
  const R = { No: 0, Shift: 1, Fleet: 2, ld: 3, tp: 4, from: 5, to: 6, lead: 7, what: 8 };
  const N = { No: 0, why: 1, Shift: 2, Fleet: 3, ld: 4, tp: 5, from: 6, to: 7, lead: 8 };
  const C = { shift: 0, fleet: 1, reg: 2, trips: 3, lead: 4, sl: 5, noGps: 6, pts: 7, slab: 8 };
  const jrKey = (s, f) => s + "|" + f;

  /* edits for one day: { jrDefault: "reported"|"low" (from the all-days setting, see editsFor), justified: true|false, cap: "shift"|"day"|"none" (none: the trips reported, as they are), jr: {"A|AV 58": {slab, just, page}}, trips: {"No": {count: true|false, slab, remark}} } */
  function computeDay(D, E) {
    E = E || {};
    const te = E.trips || {};
    // counted trips: the GPS hauls accepted, less those set "not counted", plus review / not-counted hauls set "count"
    const counted = [];
    for (const t of D.trips) {
      const e = te[t[T.No]];
      if (e && e.count === false) continue;
      counted.push({ No: t[T.No], Shift: t[T.Shift], Fleet: t[T.Fleet], ld: t[T.ld], tp: t[T.tp], from: t[T.from], to: t[T.to], lead: t[T.lead],
        oneWay: t[T.oneWay], portal: t[T.portal], pieces: [t[T.a], t[T.b], t[T.M], t[T.c], t[T.d]], note: t[T.note] || "",
        slab: (e && e.slab) || t[T.slab], edited: !!(e && (e.slab || e.remark)), remark: (e && e.remark) || "", src: "gps" });
    }
    const addBack = (rows, I, src) => {
      for (const r of rows) {
        const e = te[r[I.No]];
        if (!(e && e.count === true)) continue;
        const lead = r[I.lead];
        counted.push({ No: r[I.No], Shift: r[I.Shift], Fleet: r[I.Fleet], ld: r[I.ld], tp: r[I.tp], from: r[I.from], to: r[I.to], lead,
          oneWay: null, portal: null, pieces: [null, null, null, null, null], note: (src === "review" ? "Counted on review" : "Counted on review (was not counted)"),
          slab: e.slab || slab(lead) || "0-1", edited: true, remark: e.remark || "", src });
      }
    };
    addBack(D.review, R, "review"); addBack(D.notCounted, N, "notCounted");
    counted.sort((x, y) => (x.Shift < y.Shift ? -1 : x.Shift > y.Shift ? 1 : x.Fleet < y.Fleet ? -1 : x.Fleet > y.Fleet ? 1 : x.ld - y.ld));
    const reviewLeft = D.review.filter(r => !(te[r[R.No]] && te[r[R.No]].count === true));
    const ncLeft = D.notCounted.filter(r => !(te[r[N.No]] && te[r[N.No]].count === true));

    const gps = {}; SHIFTS.forEach(s => (gps[s] = zero()));
    const byTS = new Map();
    for (const t of counted) {
      gps[t.Shift][SL.indexOf(t.slab)] += 1;
      const k = t.Fleet + "|" + t.Shift; if (!byTS.has(k)) byTS.set(k, []); byTS.get(k).push(t);
    }
    const sum = a => a.reduce((x, y) => x + y, 0);
    const repShift = {}; SHIFTS.forEach(s => (repShift[s] = sum(D.con.filter(r => r[C.shift] === s).map(r => r[C.trips]))));
    const repDay = sum(SHIFTS.map(s => repShift[s]));
    const gpsShift = {}; SHIFTS.forEach(s => (gpsShift[s] = sum(gps[s])));
    const gpsDay = sum(SHIFTS.map(s => gpsShift[s]));

    // Justification Register: tipper-shifts with no GPS data (as found by the pipeline), in shift order A → B → C; the day cap is filled in this order
    // cap per shift (default, Aj 16.11: a shift never gets more than reported − its GPS hauls) or per day (the earlier rule); edit "cap"
    const on = E.justified !== false;
    const capBy = E.cap || D.capDefault || "shift";
    const capShift = {}; SHIFTS.forEach(s => (capShift[s] = Math.max(0, repShift[s] - gpsShift[s])));
    const capDay = Math.max(0, repDay - gpsDay);
    const cap = capBy === "none" ? null : capBy === "day" ? capDay : Math.min(capDay, sum(SHIFTS.map(s => capShift[s])));   // per shift: the day limit holds as well
    const jre = E.jr || {};
    // lead of trips given on justification (Aj 19.11.2026): "reported" = the slab of the contractor's reported lead for the tipper-shift
    // (over 5 km in 4-5), "low" = 0-1 km; a row's own choice on the Justification Register wins.
    const jrAuto = r => (E.jrDefault === "low" ? "0-1" : slab(r[C.lead]) || "0-1");
    let used = 0; const usedS = { A: 0, B: 0, C: 0 };
    const register = D.con.filter(r => r[C.noGps]).map((r, i) => {
      const k = jrKey(r[C.shift], r[C.fleet]); const e = jre[k] || {};
      const g = (byTS.get(r[C.fleet] + "|" + r[C.shift]) || []).length;
      const room = capBy === "none" ? Infinity : capBy === "day" ? capDay - used : Math.min(capShift[r[C.shift]] - usedS[r[C.shift]], capDay - used);
      const inc = on ? Math.max(0, Math.min(r[C.trips], room)) : 0; used += inc; usedS[r[C.shift]] += inc;
      return { sl: i + 1, key: k, shift: r[C.shift], fleet: r[C.fleet], reg: r[C.reg], reported: r[C.trips], gps: g, diff: r[C.trips] - g, included: inc,
        auto: jrAuto(r), slab: e.slab || jrAuto(r), reason: `No GPS data from the device for the shift (${r[C.pts]} valid GPS points) — technical`, just: e.just || "", page: e.page || "" };
    });
    const just = {}; SHIFTS.forEach(s => (just[s] = sum(register.filter(r => r.shift === s).map(r => r.included))));

    const vtsShift = {}; SHIFTS.forEach(s => (vtsShift[s] = sum(D.vts[s])));
    const revShift = {}; SHIFTS.forEach(s => (revShift[s] = reviewLeft.filter(r => r[R.Shift] === s).length));
    const A = SHIFTS.map(s => ({ shift: s, reported: repShift[s], vts: vtsShift[s], gps: gpsShift[s], review: revShift[s],
      noGpsRep: sum(D.con.filter(r => r[C.shift] === s && r[C.noGps]).map(r => r[C.trips])), justified: just[s], final: gpsShift[s] + just[s], diff: repShift[s] - gpsShift[s] - just[s] }));
    const tot = keys => { const o = { shift: "DAY TOTAL" }; keys.forEach(k => (o[k] = sum(A.map(r => r[k])))); return o; };
    const Atot = tot(["reported", "vts", "gps", "review", "noGpsRep", "justified", "final", "diff"]);

    const B = SHIFTS.map(s => {
      const v = gps[s].slice();
      register.filter(r => r.shift === s).forEach(r => (v[SL.indexOf(r.slab)] += r.included));
      return { shift: s, slabs: v, total: sum(v), ok: sum(v) === A[SHIFTS.indexOf(s)].final };
    });
    const Bday = SL.map((_, i) => sum(B.map(r => r.slabs[i])));
    const Btot = { slabs: Bday, total: sum(Bday), ok: sum(Bday) === Atot.final };

    // undesignated (HQ NIT Lead Report): null where the day's Trip List export had DESIGNATED rows only
    const pctU = (d, u) => (u == null ? null : d + u > 0 ? u / (d + u) : 0);
    const Cv = SHIFTS.map(s => { const u = D.vtsUnd ? D.vtsUnd[s] : null; return { shift: s, slabs: D.vts[s].slice(), total: vtsShift[s], und: u, pctUnd: pctU(vtsShift[s], u) }; });
    const Cday = SL.map((_, i) => sum(Cv.map(r => r.slabs[i])));
    const undDay = D.vtsUnd ? sum(Cv.map(r => r.und)) : null;

    const cons = {}; SHIFTS.forEach(s => (cons[s] = zero()));
    for (const r of D.con) cons[r[C.shift]][SL.indexOf(r[C.slab] || "0-1")] += r[C.trips];
    const Dd = SHIFTS.map((s, i) => ({ shift: s, con: cons[s], fin: B[i].slabs, diff: B[i].slabs.map((x, k) => x - cons[s][k]) }));
    const Dtot = { con: SL.map((_, k) => sum(Dd.map(r => r.con[k]))), fin: Bday.slice() }; Dtot.diff = Dtot.fin.map((x, k) => x - Dtot.con[k]);

    // Tipper-wise
    const short = a => {
      const c = new Map();
      a.forEach(t => (t.note || "").split("; ").filter(Boolean).forEach(x => c.set(x, (c.get(x) || 0) + 1)));
      return [...c.entries()].sort((x, y) => y[1] - x[1]).slice(0, 3).map(([k, v]) => `${k} ${v}`).join("; ");
    };
    const median = a => { if (!a.length) return null; const b = a.slice().sort((x, y) => x - y); const m = b.length >> 1; return b.length % 2 ? b[m] : (b[m - 1] + b[m]) / 2; };
    const tw = [];
    const seen = new Set();
    for (const r of D.con) {
      const a = byTS.get(r[C.fleet] + "|" + r[C.shift]) || []; seen.add(r[C.fleet] + "|" + r[C.shift]);
      const by = {}; a.forEach(t => (by[t.slab] = (by[t.slab] || 0) + 1));
      const cs = r[C.slab]; let chk, rem;
      if (r[C.noGps]) { chk = "No GPS"; rem = "No GPS data: see Justification Register"; }
      else if (!a.length) { chk = "No GPS trip"; rem = ""; }
      else {
        const mx = Math.max(...Object.values(by)); const top = (by[cs] || 0) === mx ? cs : SL.find(x => (by[x] || 0) === mx);   // a tie: the contractor's slab if among the top, else the lowest
        chk = Object.keys(by).length === 1 && top === cs ? "Same" : top === cs ? "Mostly same" : SL.indexOf(top) < SL.indexOf(cs) ? "Ours lower" : "Ours higher"; rem = short(a);
      }
      tw.push({ shift: r[C.shift], fleet: r[C.fleet], rep: r[C.trips], repLead: r[C.lead], repSlab: cs, gps: a.length, bySlab: SL.filter(x => by[x]).map(x => `${x}: ${by[x]}`).join(" · "),
        gpsLead: a.length ? median(a.map(t => t.lead).filter(x => x !== null)) : null, diff: r[C.noGps] ? 0 : r[C.trips] - a.length, chk, rem });
    }
    for (const [k, a] of [...byTS].sort((x, y) => { const [f1, s1] = x[0].split("|"), [f2, s2] = y[0].split("|"); return f1 < f2 ? -1 : f1 > f2 ? 1 : s1 < s2 ? -1 : s1 > s2 ? 1 : 0; })) {   // fleet, then shift: the workbook's order
      if (seen.has(k)) continue;
      const [f, s] = k.split("|"); const by = {}; a.forEach(t => (by[t.slab] = (by[t.slab] || 0) + 1));
      tw.push({ shift: s, fleet: f, rep: 0, repLead: null, repSlab: null, gps: a.length, bySlab: SL.filter(x => by[x]).map(x => `${x}: ${by[x]}`).join(" · "),
        gpsLead: median(a.map(t => t.lead).filter(x => x !== null)), diff: -a.length, chk: "Not reported", rem: short(a) });
    }
    const forReview = counted.filter(t => t.lead > 5 && t.src === "gps").map(t => ({ ...t, why: "Lead over 5 km", countedTxt: "Yes, in 4-5" }))
      .concat(reviewLeft.map(r => ({ No: r[R.No], Shift: r[R.Shift], Fleet: r[R.Fleet], ld: r[R.ld], tp: r[R.tp], from: r[R.from], to: r[R.to], lead: r[R.lead], why: r[R.what], countedTxt: "No" })));
    return { day: D.day, counted, register, cap, capBy, capShift, justifiedOn: on, A, Atot, B, Btot, C: Cv, Ctot: { slabs: Cday, total: sum(Cday), und: undDay, pctUnd: pctU(sum(Cday), undDay) }, D: Dd, Dtot, tw, forReview, notCounted: ncLeft };
  }
  // ---------- contractor's Trip Reconciliation format (Ajantha, 18.11.2026): slabs 0-1 … 8-9 km, per shift + day total.
  // Designated / Undesignated: the portal's trips on the portal's own lead (VTS = nitReport), shift by start time; Contractor Reported: each
  // tipper-shift in the slab of its reported lead; % Deviation w.r.t Reported: the contractor's formula (Reported − Total) × 100 ÷ Total,
  // "–" where Total is 0; Final after Reconciliation: reconFinal below; GPS hauls: internal only. Remarks say how VTS trips turn into the final.
  // Aj 21.11.2026: slabs 5-6 … 8-9 removed — the NIT's slabs only, any trip over 5 km in 4-5
  const RSL = ["0-1", "1-2", "2-3", "3-4", "4-5"];
  const RECON_HEAD = { mine: "Hingula OCP, Hingula Area", nit: "NIT-185", contractor: "Ajantha BLAPPL(JV)" };   // Aj 21.11.2026
  const fold = a => RSL.map((_, i) => (i < RSL.length - 1 ? (a && a[i]) || 0 : ((a && a.slice(i)) || []).reduce((x, y) => x + (y || 0), 0)));
  const norm = x => ({ des: fold(x.des), und: fold(x.und), rep: fold(x.rep), gps: fold(x.gps), rem: x.rem || {} });
  const sum = a => a.reduce((x, y) => x + y, 0);
  const pad = fold;
  const dev = (rep, tot) => (tot ? Math.round((rep - tot) * 100 / tot * 100) / 100 : null);
  function reconBlock(label, x, fin, just) {
    const tot = x.des.map((v, i) => v + x.und[i]);
    return { label, des: x.des, und: x.und, tot, rep: x.rep, dev: x.rep.map((v, i) => dev(v, tot[i])), fin: pad(fin), gps: x.gps, rem: Object.assign({ just }, x.rem),
      sums: { des: sum(x.des), und: sum(x.und), tot: sum(tot), rep: sum(x.rep), fin: sum(fin), gps: sum(x.gps) } };
  }
  function addRec(a, b) { const o = {}; for (const k of ["des", "und", "rep", "gps"]) o[k] = a[k].map((v, i) => v + b[k][i]); o.rem = {}; for (const k in b.rem) o.rem[k] = (a.rem[k] || 0) + b.rem[k]; return o; }
  const emptyRec = () => ({ des: RSL.map(() => 0), und: RSL.map(() => 0), rep: RSL.map(() => 0), gps: RSL.map(() => 0), rem: {} });
  // Final after Reconciliation (Aj 19.11.2026 12:10, supersedes 02:21): the same figures as the SOP Daily Report — GPS hauls on the haul lead
  // plus trips given on justification (all-days switch / per-day switch, lead per the all-days setting), slabs 0-1 … 4-5 (NIT's highest).
  const rslab = x => (x === null || x === undefined || Number.isNaN(x) ? 0 : x <= 1 ? 0 : Math.min(RSL.length - 1, Math.ceil(x) - 1));
  function reconFinal(D, r) {   // = the SOP Daily Report's final by slab (Aj 19.11.2026 12:10): 0-1 … 4-5, over 5 km in 4-5
    const f = {}; SHIFTS.forEach((s, i) => (f[s] = pad(r.B[i].slabs)));
    return f;
  }
  function reconDay(D, r) {
    if (!D.recon) return null;
    const f = reconFinal(D, r);
    const blocks = SHIFTS.map((s, i) => reconBlock(s, norm(D.recon.shifts[s]), f[s], r.A[i].justified));
    const t = SHIFTS.reduce((a, s) => addRec(a, norm(D.recon.shifts[s])), emptyRec());
    blocks.push(reconBlock("Day", t, RSL.map((_, i) => sum(SHIFTS.map(s => f[s][i]))), r.Atot.justified));
    return blocks;
  }
  function reconPeriod(list) {   // list of [D, r]: one total block over the days
    const ok = list.filter(([D]) => D.recon); if (!ok.length) return null;
    let t = emptyRec(), fin = RSL.map(() => 0), just = 0;
    for (const [D, r] of ok) { t = SHIFTS.reduce((a, s) => addRec(a, norm(D.recon.shifts[s])), t); const f = reconFinal(D, r); SHIFTS.forEach(s => f[s].forEach((v, i) => (fin[i] += v))); just += r.Atot.justified; }
    return reconBlock("Total", t, fin, just);
  }
  function reconRemark(b) {
    const m = b.rem, p = [];
    if (m.raised) p.push(`${m.raised} trip(s) cut at a terminal boundary (loading or tipping outside the VTS terminal): placed in the slab of the full haul`);
    if (m.lowered) p.push(`${m.lowered} trip(s) in a lower slab (lead measured to the actual loading / tip place)`);
    if (m.split) p.push(`${m.split} VTS trip(s) were part of a haul already counted (haul split in two): counted once`);
    if (m.notHaul) p.push(`${m.notHaul} VTS trip(s) (${m.notHaulUnd || 0} undesignated) are short moves, not hauls: not counted`);
    if (m.noVts) p.push(`${m.noVts} haul(s) with no VTS trip: added`);
    if (m.just) p.push(`${m.just} trip(s) given on justification (no GPS data: GPS not working / under maintenance): see the Justification Register`);
    p.push(`Reported − Final: ${b.sums.rep - b.sums.fin}`);
    return p.join("; ");
  }
  // all-days settings (Aj 19.11.2026): edits.global = { justified: false (switch off for every day), jrLead: "reported" (default) | "low" }.
  // A day's own switch, when set, wins over the all-days one.
  function editsFor(edits, day) {
    const G = (edits && edits.global) || {}, d = ((edits && edits.days) || {})[day] || {};
    const o = Object.assign({}, d, { jrDefault: G.jrLead || "reported" });
    if (d.justified === undefined && G.justified === false) o.justified = false;
    return o;
  }
  // ---------- tipper-wise: contractor reported against VTS trips, GPS hauls and the final (sheet "Over-reporting", Aj 19.11.2026)
  function tipperRows(D, r) {
    const key = (s, f) => s + "|" + f, M = new Map();
    const get = (s, f) => { const k = key(s, f); if (!M.has(k)) M.set(k, { shift: s, fleet: f, reg: "", rep: 0, lead: null, des: 0, und: 0, gps: 0, just: 0 }); return M.get(k); };
    for (const c of D.con) { const x = get(c[C.shift], c[C.fleet]); x.rep += c[C.trips]; x.lead = c[C.lead]; x.reg = c[C.reg] || x.reg; }
    for (const t of r.counted) get(t.Shift, t.Fleet).gps += 1;
    for (const g of r.register) get(g.shift, g.fleet).just += g.included;
    const ts = (D.recon && D.recon.tippers) || {};
    for (const k in ts) { const [s, f] = k.split("|"); const x = get(s, f); x.des += ts[k][0]; x.und += ts[k][1]; if (!x.reg) x.reg = ts[k][2] || ""; }
    const rows = [...M.values()].map(x => Object.assign(x, { fin: x.gps + x.just, overGps: x.rep - x.gps, overFin: x.rep - x.gps - x.just }));
    rows.sort((a, b) => (a.shift < b.shift ? -1 : a.shift > b.shift ? 1 : a.fleet.localeCompare(b.fleet, "en", { numeric: true })));
    return rows;
  }
  // ---------- Clusters summary over the days shown (Aj 19.11.2026): patch → dump location by day, totals by day, position by day
  const havM = (a, b) => { const r = Math.PI / 180, x = Math.sin((b[0] - a[0]) * r / 2) ** 2 + Math.cos(a[0] * r) * Math.cos(b[0] * r) * Math.sin((b[1] - a[1]) * r / 2) ** 2; return 2 * 6371000 * Math.asin(Math.sqrt(x)); };
  function clusterSummary(days) {
    const ds = days.filter(D => D.clusters).map(D => D.day), idx = new Map(ds.map((d, i) => [d, i])), z = () => ds.map(() => 0);
    const route = new Map(), patch = new Map(), dump = new Map(), posL = new Map(), posT = new Map();
    const bump = (M, k, i, n) => { if (!M.has(k)) M.set(k, z()); M.get(k)[i] += n; };
    for (const D of days) { if (!D.clusters) continue; const i = idx.get(D.day);
      for (const r of D.clusters.routes) { bump(route, (r.P || "?") + " → " + (r.Q || "?"), i, r.hauls); bump(patch, r.P || "?", i, r.hauls); bump(dump, r.Q || "?", i, r.hauls); }
      for (const [M, kind] of [[posL, "loading"], [posT, "tipping"]]) for (const [k, v] of Object.entries((D.clusters.pos || {})[kind] || {})) {
        if (k.startsWith("Scattered")) continue; if (!M.has(k)) M.set(k, ds.map(() => null)); M.get(k)[i] = v; } }
    const rows = M => [...M.entries()].map(([k, a]) => ({ name: k, byDay: a, total: a.reduce((x, y) => x + y, 0) })).sort((a, b) => b.total - a.total);
    const pos = M => [...M.entries()].map(([k, a]) => { const f = a.find(v => v); return { name: k, byDay: a.map(v => (v ? { lat: v[0], lng: v[1], hauls: v[2], moved: Math.round(havM(f, v)) } : null)) }; });
    return { days: ds, routes: rows(route), patches: rows(patch), dumps: rows(dump), posL: pos(posL), posT: pos(posT) };
  }
  const api = { SL, SH, SHIFTS, slab, computeDay, jrKey, T, R, N, C, RSL, RECON_HEAD, reconDay, reconPeriod, reconRemark, editsFor, tipperRows, clusterSummary };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.SOPCompute = api;
})(typeof self !== "undefined" ? self : this);
