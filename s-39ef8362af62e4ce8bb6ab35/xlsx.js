/* Excel download of the SOP for one day or several days (in date order), from computeDay results.
   buildWorkbook(ExcelJS, data, days, edits) → ExcelJS.Workbook. Same module in the page and in node. */
(function (root) {
  "use strict";
  const S = (typeof module !== "undefined" && module.exports) ? require("./compute.js") : root.SOPCompute;
  const { SL, SH, SHIFTS } = S;
  const DEFAULT_SIGN = [
    { role: "Contractor / Representative", org: "Ajantha BLAPPL(JV)" },
    { role: "Control Room In-charge", org: "Hingula OCP" },
    { role: "OB In-charge", org: "Hingula OCP" },
    { role: "Mine Manager", org: "Hingula OCP" },
  ];
  // reconciliation sheet (Aj 19.11.2026): the SOP's four, then HOD (Survey) and Project Officer
  const RECON_EXTRA = [{ role: "HOD (Survey)", org: "Hingula OCP" }, { role: "Project Officer", org: "Hingula OCP" }];
  const reconSign = edits => [...((edits && edits.sign && edits.sign.blocks) || DEFAULT_SIGN), ...RECON_EXTRA];
  const dmy = d => d.slice(8, 10) + "." + d.slice(5, 7) + "." + d.slice(0, 4);
  const ist = ms => { const d = new Date(ms + 19800000); const p = n => String(n).padStart(2, "0");
    return `${p(d.getUTCDate())}.${p(d.getUTCMonth() + 1)}.${d.getUTCFullYear()} ${p(d.getUTCHours())}:${p(d.getUTCMinutes())}:${p(d.getUTCSeconds())}`; };
  const thin = { style: "thin", color: { argb: "FF9AA5B1" } };
  const BOX = { top: thin, left: thin, bottom: thin, right: thin };
  const HEAD = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDE7F3" } };
  const TOT = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1EFE8" } };
  const YEL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
  const RED = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8D7D3" } };

  function grid(ws, r0, c0, r1, c1, total) {
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
      const x = ws.getCell(r, c); x.border = BOX;
      if (r === r0) { x.font = { bold: true }; x.fill = HEAD; x.alignment = { wrapText: true, vertical: "middle", horizontal: "center" }; }
      else if (total && r === r1) { x.font = { bold: true }; x.fill = TOT; }
    }
  }
  const add = (ws, vals) => ws.addRow(vals).number;
  function line(ws, text, opt) {
    const r = add(ws, [text]); const c = ws.getCell(r, 1);
    if (opt && opt.bold) c.font = { bold: true, size: opt.size || 11 };
    if (opt && opt.merge) { ws.mergeCells(r, 1, r, opt.merge); c.alignment = { wrapText: true, vertical: "top" }; ws.getRow(r).height = opt.h || 30; }
    return r;
  }
  function signatures(ws, r, ncol, blocks) {
    const n = blocks.length; if (!n) return r;
    const w = Math.max(1, Math.floor(ncol / n));
    blocks.forEach((b, k) => {
      const c0 = 1 + k * w, c1 = c0 + w - 1;
      [["Signature", false], ["Name: ____________________", false], [b.role, true], [b.org, true], ["Date: ____________", false]].forEach(([t, bold], i) => {
        if (c1 > c0) ws.mergeCells(r + i, c0, r + i, c1);
        const x = ws.getCell(r + i, c0); x.value = t; x.alignment = { horizontal: "center" }; x.font = { bold };
        if (i === 0) for (let c = c0; c <= c1; c++) ws.getCell(r, c).border = { top: thin };
      });
    });
    return r + 5;
  }
  function landscape(ws, onePage, repeat) {
    ws.pageSetup = { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: onePage ? 1 : 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };
    if (repeat) ws.pageSetup.printTitlesRow = repeat;
  }

  function sopSheet(wb, name, D, r, data, sign) {
    const ws = wb.addWorksheet(name);
    const day = dmy(D.day);
    line(ws, `MAHANADI COALFIELDS LIMITED — ${data.site}`, { bold: true, size: 13 });
    line(ws, `DAILY / SHIFT-WISE RECONCILIATION OF OB TIPPER TRIPS — ${data.nit} ${data.contractor}`, { bold: true });
    line(ws, "Trips and lead-wise trips from the GPS track");
    line(ws, `Date: ${day}   (report day 05:00 → next day 05:00)`);
    line(ws, "Final trips = GPS hauls accepted + trips given for tipper-shifts with no GPS data. Lead = ½ inside the loading terminal + between + ½ inside the unloading terminal (device odometer), at the real loading and tip places.");
    add(ws, []);
    const rs = add(ws, ["Include trips given — justified (Yes / No):", r.justifiedOn ? "Yes" : "No", "← justified = trips of tipper-shifts with no GPS data (sheet Justification Register), given only up to (reported − GPS hauls) for the day; slab per row, default 0–1"]);
    ws.getCell(rs, 1).font = { bold: true }; ws.getCell(rs, 2).fill = YEL; ws.getCell(rs, 2).border = BOX;
    const rc = add(ws, ["Cap on justified trips:", r.capBy === "none" ? "No cap" : r.capBy === "shift" ? "Per shift" : "Per day", r.capBy === "none" ? "← the trips reported, as they are" : r.capBy === "shift" ? `← each shift: reported − GPS hauls of the shift (A ${r.capShift.A} · B ${r.capShift.B} · C ${r.capShift.C}), day ${r.cap} at most` : `← the day: reported − GPS hauls (${r.cap}), filled A → B → C`]);
    ws.getCell(rc, 1).font = { bold: true }; ws.getCell(rc, 2).border = BOX;
    add(ws, []);
    line(ws, "A.  Trips: contractor, VTS and GPS", { bold: true });
    let r0 = add(ws, ["Shift", "Trips reported by contractor", "Trips generated by VTS (designated)", "GPS hauls accepted", "GPS hauls under review (not counted)", "No-GPS trips reported", "Trips given — justified", "FINAL TRIPS", "Difference (Reported − Final)"]);
    ws.getRow(r0).height = 62;
    r.A.forEach(a => add(ws, [SH[a.shift], a.reported, a.vts, a.gps, a.review, a.noGpsRep, a.justified, a.final, a.diff]));
    const t = r.Atot; let r1 = add(ws, ["DAY TOTAL", t.reported, t.vts, t.gps, t.review, t.noGpsRep, t.justified, t.final, t.diff]); grid(ws, r0, 1, r1, 9, true);
    add(ws, []);
    line(ws, "B.  Final trips by lead slab (km): lower < lead ≤ upper. A lead over 5 km is counted in 4–5 (highest slab of the NIT) and listed in For review. Trips given — justified are in the slab chosen for each row of the Justification Register (default 0–1).", { bold: true });
    r0 = add(ws, ["Shift", ...SL, "Total", "Total = Final?"]);
    r.B.forEach(b => { const q = add(ws, [SH[b.shift], ...b.slabs, b.total, b.ok ? "OK" : "MISMATCH"]); if (!b.ok) ws.getCell(q, 8).fill = RED; });
    r1 = add(ws, ["DAY TOTAL", ...r.Btot.slabs, r.Btot.total, r.Btot.ok ? "OK" : "MISMATCH"]); grid(ws, r0, 1, r1, 8, true);
    add(ws, []);
    line(ws, "C.  For comparison: VTS designated trips by VTS distance slab (km), and undesignated", { bold: true });
    r0 = add(ws, ["Shift", ...SL, "Total", "VTS undesignated", "% undesignated (of designated + undesignated)"]); ws.getRow(r0).height = 45;
    const uC = (u, p) => [u == null ? "—" : u, p == null ? "—" : p];
    r.C.forEach(c => add(ws, [SH[c.shift], ...c.slabs, c.total, ...uC(c.und, c.pctUnd)]));
    r1 = add(ws, ["DAY TOTAL", ...r.Ctot.slabs, r.Ctot.total, ...uC(r.Ctot.und, r.Ctot.pctUnd)]); grid(ws, r0, 1, r1, 9, true);
    for (let q = r0 + 1; q <= r1; q++) ws.getCell(q, 9).numFmt = "0.0%";
    line(ws, "VTS designated and undesignated trips of NIT-185 (Transporter " + data.contractor + ") from the VTS Trip List; % undesignated as in HQ's NIT Lead Report." + (r.Ctot.und == null ? " This day's Trip List export has DESIGNATED rows only, so undesignated is not known (—)." : ""));
    add(ws, []);
    line(ws, "D.  Lead-wise: contractor's reported trips against final trips", { bold: true });
    r0 = add(ws, ["Shift", ...SL.map(x => `${x} contractor`), ...SL.map(x => `${x} final`), ...SL.map(x => `${x} final − contractor`)]); ws.getRow(r0).height = 62;
    r.D.forEach(d => add(ws, [SH[d.shift], ...d.con, ...d.fin, ...d.diff]));
    r1 = add(ws, ["DAY TOTAL", ...r.Dtot.con, ...r.Dtot.fin, ...r.Dtot.diff]); grid(ws, r0, 1, r1, 16, true);
    add(ws, []);
    if (sign.show) {
      const who = sign.blocks.map(b => b.role).join(", ");
      line(ws, `Every trip given on justification is listed in the Justification Register with its reason, slab and bound register page, for the counter-signature of ${who} (SOP 7).`, { merge: 16 });
      line(ws, "Certified that the shift-wise reconciliation of OB tipper trips with the trips reported by the contractor has been carried out jointly from the GPS track of each tipper, and that no trip beyond the GPS hauls has been taken except as justified in the bound-page register.", { merge: 16 });
      line(ws, "Reconciled jointly by", { bold: true }); add(ws, []); add(ws, []);
      signatures(ws, ws.rowCount + 1, 16, sign.blocks);
    }
    ws.getColumn(1).width = 18; for (let c = 2; c <= 16; c++) ws.getColumn(c).width = 12; ws.getColumn(3).width = 14;
    landscape(ws, true);
  }

  function justSheet(wb, name, D, r, data, sign) {
    const ws = wb.addWorksheet(name);
    const day = dmy(D.day);
    [[`MAHANADI COALFIELDS LIMITED — ${data.site}`, 13], [`JUSTIFICATION OF MISSING / UNACCOUNTED TRIPS — ${day}  (SOP 7: to be counter-signed)`, 11],
     ["SOP: Updation of Loading & Un-loading Points in VTS Portal for OB Tippers and OB Reporting (MCL HQ, Production)", 0],
     [`Date: ${day}   (report day 05:00 → next day 05:00)        NIT: ${data.nit}        Contractor: ${data.contractor}`, 0]].forEach(([t, sz]) => {
      const q = add(ws, [t]); ws.mergeCells(q, 1, q, 13 + (sign.show ? sign.blocks.length : 0)); if (sz) ws.getCell(q, 1).font = { bold: true, size: sz };
    });
    add(ws, []);
    const sigCols = sign.show ? sign.blocks.map(b => b.role) : [];
    const ncol = 13 + sigCols.length;
    const r0 = add(ws, ["Sl.", "Date", "Shift", "Fleet No", "Reg No", "Reported", "GPS hauls", "Difference", "Trips included (+)", "Slab", "Reason", "Justification", "Bound register page", ...sigCols]);
    ws.getRow(r0).height = 45;
    let r1 = r0;
    if (!r.register.length) { r1 = add(ws, ["No trip was given on justification on this day."]); ws.mergeCells(r1, 1, r1, ncol); }
    else {
      r.register.forEach(j => { const q = add(ws, [j.sl, day, j.shift, j.fleet, j.reg, j.reported, j.gps, j.diff, j.included, j.slab, j.reason, j.just, j.page]);
        ws.getCell(q, 10).fill = YEL; ws.getCell(q, 10).alignment = { horizontal: "center" }; [11, 12].forEach(c => (ws.getCell(q, c).alignment = { wrapText: true, vertical: "middle" })); });
      const s = k => r.register.reduce((a, j) => a + j[k], 0);
      r1 = add(ws, ["TOTAL", null, null, null, null, s("reported"), s("gps"), s("diff"), s("included")]);
    }
    grid(ws, r0, 1, r1, ncol, r.register.length > 0);
    add(ws, []);
    line(ws, `Only mismatches arising from technical reasons may be included (SOP 7). Trips included are ` + (r.capBy === "none" ? "the trips reported, as they are (no cap)." : r.capBy === "shift" ? `capped at each shift's reported trips − GPS hauls (A ${r.capShift.A} · B ${r.capShift.B} · C ${r.capShift.C}), never above the day's (${r.cap}).` : `capped at the day's reported trips − GPS hauls (${r.cap} on this day), filled shift A, then B, then C.`) + (r.justifiedOn ? "" : " Trips given — justified are switched OFF for this day."), { merge: ncol });
    line(ws, "Slab: 0–1 unless the joint check finds otherwise. Each entry must be written in the bound-page register and counter-signed there; this sheet is the list to be signed against.", { merge: ncol });
    if (sign.show) { add(ws, []); add(ws, []); signatures(ws, ws.rowCount + 1, 16, sign.blocks); }
    [5, 11, 6, 9, 13, 9, 9, 10, 10, 8, 34, 30, 11, ...sigCols.map(() => 12)].forEach((w, i) => (ws.getColumn(i + 1).width = w));
    ws.views = [{ state: "frozen", ySplit: r0 }];
    landscape(ws, false, `${r0}:${r0}`);
  }

  function table(wb, name, head, rows, widths, freeze) {
    const ws = wb.addWorksheet(name); add(ws, head); rows.forEach(x => ws.addRow(x)); grid(ws, 1, 1, 1, head.length, false);
    widths.forEach((w, i) => (ws.getColumn(i + 1).width = w)); ws.views = [{ state: "frozen", ySplit: 1, xSplit: freeze || 0 }]; ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: head.length } };
    return ws;
  }

  function buildWorkbook(ExcelJS, data, days, edits) {
    edits = edits || {}; const sign = { show: !(edits.sign && edits.sign.show === false), blocks: (edits.sign && edits.sign.blocks) || DEFAULT_SIGN };
    const wb = new ExcelJS.Workbook(); wb.creator = "Lead-wise SOP"; wb.created = new Date();
    const res = days.map(D => [D, S.computeDay(D, S.editsFor(edits, D.day))]);
    if (res.length > 1) {
      const ws = wb.addWorksheet("Month");
      line(ws, `MAHANADI COALFIELDS LIMITED — ${data.site} — ${data.nit} ${data.contractor}`, { bold: true, size: 13 });
      line(ws, `Lead-wise reconciliation of OB tipper trips, ${dmy(res[0][0].day)} to ${dmy(res[res.length - 1][0].day)} (report day 05:00 → next day 05:00)`, { bold: true });
      add(ws, []);
      const r0 = add(ws, ["Date", "Reported by contractor", "VTS designated", "VTS undesignated", "% undesignated", "GPS hauls accepted", "Under review (not counted)", "Trips given — justified", "FINAL TRIPS", "Difference (Reported − Final)", ...SL.map(x => `Final ${x}`)]);
      ws.getRow(r0).height = 45;
      const tot = Array(7 + SL.length).fill(0); let dK = 0, uK = 0, nK = 0;
      res.forEach(([D, r]) => { const v = [r.Atot.reported, r.Atot.vts, r.Atot.gps, r.Atot.review, r.Atot.justified, r.Atot.final, r.Atot.diff, ...r.Btot.slabs]; v.forEach((x, i) => (tot[i] += x));
        const u = r.Ctot.und; if (u != null) { dK += r.Ctot.total; uK += u; nK++; }
        const row = v.slice(); row.splice(2, 0, u == null ? "—" : u, r.Ctot.pctUnd == null ? "—" : r.Ctot.pctUnd); add(ws, [dmy(D.day), ...row]); });
      const tr = tot.slice(); tr.splice(2, 0, nK ? uK : "—", nK ? (dK + uK ? uK / (dK + uK) : 0) : "—");
      const r1 = add(ws, ["TOTAL", ...tr]); grid(ws, r0, 1, r1, 15, true);
      for (let q = r0 + 1; q <= r1; q++) ws.getCell(q, 5).numFmt = "0.0%";
      line(ws, "VTS designated / undesignated: NIT-185's own trips in the VTS Trip List; % undesignated = undesignated ÷ (designated + undesignated), as in HQ's NIT Lead Report; — : export had designated trips only. TOTAL % over the days where it is known.");
      ws.getColumn(1).width = 12; for (let c = 2; c <= 15; c++) ws.getColumn(c).width = 12;
      landscape(ws, true);
    }
    for (const [D, r] of res) {
      const tag = res.length > 1 ? dmy(D.day).slice(0, 5) + " " : "";
      sopSheet(wb, tag + (res.length > 1 ? "SOP" : "SOP Daily Report"), D, r, data, sign);
      justSheet(wb, tag + (res.length > 1 ? "Justification" : "Justification Register"), D, r, data, sign);
    }
    const many = res.length > 1; const dcol = D => (many ? [dmy(D.day)] : []); const dh = many ? ["Date"] : []; const dw = many ? [11] : [];
    table(wb, "Tipper-wise", [...dh, "Shift", "Tipper", "Reported trips", "Reported lead (km)", "Reported slab", "GPS trips", "GPS trips by slab", "GPS lead, median (km)", "Trips: reported − GPS", "Slab check", "Remark", "Reviewer remark", "Contractor's reply"],
      res.flatMap(([D, r]) => r.tw.map(t => [...dcol(D), t.shift, t.fleet, t.rep, t.repLead === null ? null : +t.repLead.toFixed(2), t.repSlab, t.gps, t.bySlab, t.gpsLead === null ? null : +t.gpsLead.toFixed(2), t.diff, t.chk, t.rem, null, null])),
      [...dw, 6, 9, 9, 9, 8, 8, 26, 9, 10, 12, 46, 24, 30], dw.length + 2);
    { const ws = wb.getWorksheet("Tipper-wise"), all = res.flatMap(([, r]) => r.tw), sm = k => all.reduce((a, t) => a + (t[k] || 0), 0);   // totals (16.11)
      const o = dh.length, v = Array(o + 13).fill(null); v[0] = "TOTAL"; v[o + 2] = sm("rep"); v[o + 5] = sm("gps"); v[o + 8] = sm("diff");
      ws.addRow([]); const row = ws.addRow(v); row.eachCell({ includeEmpty: true }, c => { c.font = { bold: true }; c.fill = TOT; }); }
    table(wb, "Trips", [...dh, "No", "Shift", "Tipper", "Leaves loading stop", "Tip", "From", "To", "Lead (km)", "Slab", "One way (km)", "Portal formula (km)", "a: in, face", "b: out, face", "M: between", "c: in, dump", "d: out, dump", "Note", "Edited on the page"],
      res.flatMap(([D, r]) => r.counted.map(t => [...dcol(D), t.No, t.Shift, t.Fleet, ist(t.ld), ist(t.tp), t.from, t.to, t.lead, t.slab, t.oneWay, t.portal, ...t.pieces, t.note, t.edited ? ("yes" + (t.remark ? ": " + t.remark : "")) : null])),
      [...dw, 7, 6, 8, 19, 19, 20, 20, 9, 6, 9, 10, 8, 8, 8, 8, 8, 30, 20], dw.length + 4);
    table(wb, "For review", [...dh, "No", "Shift", "Tipper", "Leaves loading stop", "Tip", "From", "To", "Lead (km)", "Why", "Counted?"],
      res.flatMap(([D, r]) => r.forReview.map(t => [...dcol(D), t.No, t.Shift, t.Fleet, ist(t.ld), ist(t.tp), t.from, t.to, t.lead, t.why, t.countedTxt])), [...dw, 7, 6, 8, 19, 19, 20, 20, 9, 40, 11]);
    const N = S.N;
    table(wb, "Not counted", [...dh, "No", "Why", "Shift", "Tipper", "Leaves loading stop", "Arrives at the far end", "From", "To", "Odometer lead (km)"],
      res.flatMap(([D, r]) => r.notCounted.map(x => [...dcol(D), x[N.No], x[N.why], x[N.Shift], x[N.Fleet], ist(x[N.ld]), ist(x[N.tp]), x[N.from], x[N.to], x[N.lead]])), [...dw, 7, 34, 6, 8, 19, 19, 20, 20, 9], dw.length + 2);
    // Over-reporting (Aj 19.11.2026): every tipper-shift — reported against VTS trips, GPS hauls, trips given on justification and the final
    { const ws = table(wb, "Over-reporting", ["Date", "Shift", "Tipper", "Registration", "Reported trips", "Reported lead (km)", "VTS designated", "VTS undesignated", "GPS hauls", "Given — justified", "Final", "Reported − GPS hauls", "Reported − Final"],
        res.flatMap(([D, r]) => S.tipperRows(D, r).map(x => [dmy(D.day), x.shift, x.fleet, x.reg, x.rep, x.lead === null ? null : +(+x.lead).toFixed(3), x.des, x.und, x.gps, x.just, x.fin, x.overGps, x.overFin])),
        [11, 6, 9, 12, 9, 9, 9, 9, 9, 9, 9, 10, 10], 3);
      ws.eachRow((row, i) => { if (i > 1 && row.getCell(13).value > 0) row.getCell(13).fill = YEL; if (i > 1 && row.getCell(12).value > 0) row.getCell(12).font = { bold: true }; });
      const all = res.flatMap(([D, r]) => S.tipperRows(D, r)), sm = k => all.reduce((a, x) => a + x[k], 0);
      ws.addRow([]); const t = ws.addRow(["TOTAL", null, null, null, sm("rep"), null, sm("des"), sm("und"), sm("gps"), sm("just"), sm("fin"), sm("overGps"), sm("overFin")]);
      t.eachCell({ includeEmpty: true }, c => { c.font = { bold: true }; c.fill = TOT; });
      ws.addRow(["Over reported (positive differences only): " + all.filter(x => x.overFin > 0).reduce((a, x) => a + x.overFin, 0) + " trips on " + all.filter(x => x.overFin > 0).length + " tipper-shifts against the final; " + all.filter(x => x.overGps > 0).reduce((a, x) => a + x.overGps, 0) + " against GPS hauls. VTS trips by start time; GPS hauls and the final as on the SOP."]); }
    // Clusters (Aj 19.11.2026): every GPS haul's loading place and tipping place, as the Tampermonkey Actual Lead
    { const cr = res.filter(([D]) => D.clusters);
      if (cr.length) {
        const Q = S.clusterSummary(cr.map(([D]) => D)), dh = Q.days.map(d => dmy(d).slice(0, 5));
        const ws = wb.addWorksheet("Clusters — summary");
        line(ws, `GPS hauls by patch and dump location, day-wise — ${dmy(Q.days[0])} to ${dmy(Q.days[Q.days.length - 1])}`, { bold: true, size: 12 });
        line(ws, "A patch is one area of working (several faces); a dump location can hold several tipping places. Patches are kept by ground, not by terminal name (from 13.11 the single terminal \"NIT 185 OB FACE\" covers Kusumpal, Atlanta and Old WB); a patch moves with the work and keeps its name; \"New patch near …\" = a working away from every patch; \"Scattered\" = small places away from every patch.");
        const block = (title, first, rows) => { add(ws, []); line(ws, title, { bold: true }); const r0 = add(ws, [first, ...dh, "Total"]);
          rows.forEach(r => add(ws, [r.name, ...r.byDay.map(v => v || null), r.total]));
          const r1 = add(ws, ["TOTAL", ...Q.days.map((_, i) => rows.reduce((a, r) => a + r.byDay[i], 0)), rows.reduce((a, r) => a + r.total, 0)]); grid(ws, r0, 1, r1, dh.length + 2, true); };
        block("Patch → dump location", "Route", Q.routes); block("Patches (loading)", "Patch", Q.patches); block("Dump locations (tipping)", "Dump location", Q.dumps);
        const posB = (title, rows) => { add(ws, []); line(ws, title, { bold: true }); const r0 = add(ws, ["Patch / location", ...dh]);
          rows.forEach(r => add(ws, [r.name, ...r.byDay.map(v => (v ? v.moved : null))])); grid(ws, r0, 1, ws.rowCount, dh.length + 1, false); };
        posB("Patch position — metres moved from its first day shown (centres: sheet Clusters — places)", Q.posL); posB("Dump location position — metres moved from its first day shown", Q.posT);
        ws.getColumn(1).width = 58; for (let c = 2; c <= dh.length + 2; c++) ws.getColumn(c).width = 8; landscape(ws, false);
        table(wb, "Clusters — routes", ["Date", "Loading place", "Tipping place", "GPS hauls", "Shift A", "Shift B", "Shift C", "Lead median (km)", "Lead min", "Lead max", "Portal formula median (km)", ...SL.map(x => "Slab " + x), "To check on the portal: tipper (registration), leaves loading → tip"],
          cr.flatMap(([D]) => D.clusters.routes.map(x => [dmy(D.day), x.L + " [" + (x.P || "") + "]", x.D + " [" + (x.Q || "") + "]", x.hauls, ...x.sh, x.med, x.min, x.max, x.portal, ...x.slabs, (x.ex || []).join("; ")])), [11, 30, 30, 9, 7, 7, 7, 9, 8, 8, 10, 7, 7, 7, 7, 7, 90], 3);
        table(wb, "Clusters — places", ["Date", "Kind", "Place", "Patch / dump location", "Latitude", "Longitude", "GPS hauls", "Shift A", "Shift B", "Shift C", "To check on the portal: tipper (registration), leaves loading → tip"],
          cr.flatMap(([D]) => D.clusters.places.map(x => [dmy(D.day), x.kind, x.name, x.patch || "", x.lat, x.lng, x.hauls, ...x.sh, (x.ex || []).join("; ")])), [11, 9, 32, 44, 10, 10, 9, 7, 7, 7, 90], 3);
      } }
    const log = (edits.log || []).filter(l => res.some(([D]) => D.day === l.day || !l.day));
    table(wb, "Changes", ["When", "Date", "Change"], log.map(l => [l.t, l.day ? dmy(l.day) : "all days", l.what]), [18, 11, 90]);
    const M = wb.addWorksheet("Method"); (data.method || "").split("\n").forEach(x => M.addRow([x]));
    res.forEach(([D]) => { M.addRow([]); M.addRow([`Terminals of ${dmy(D.day)} (geofence ids on the GPS points)`]).font = { bold: true }; D.terms.forEach(x => M.addRow([x])); });
    M.getColumn(1).width = 140;
    return wb;
  }
  // the contractor's Trip Reconciliation format (Ajantha, 18.11.2026), one sheet: per day shift A, B, C and the day, then the period total.
  // Total, % Deviation (their formula, wrapped in IFERROR) are Excel formulas; GPS hauls only when opts.gps (internal — not for sending).
  function buildReconWorkbook(ExcelJS, data, days, edits, opts) {
    opts = opts || {}; edits = edits || {};
    const wb = new ExcelJS.Workbook(); wb.creator = "Lead-wise SOP"; wb.created = new Date();
    const ws = wb.addWorksheet("Trip Reconciliation"), RSL = S.RSL, nS = RSL.length, cS = 6, cR = cS + nS;   // F … N slabs, O remarks
    const res = days.map(D => [D, S.computeDay(D, S.editsFor(edits, D.day))]).filter(([D]) => D.recon);
    add(ws, ["Name of the Mine", "NIT-No. ", "Contractor", "Date/Shift", "Trips", "Lead Distance (in Km)", ...Array(nS - 1).fill(null), "Remarks"]);
    add(ws, [null, null, null, null, null, ...RSL, null]);
    ws.mergeCells(1, cS, 1, cS + nS - 1); [1, 2, 3, 4, 5, cR].forEach(c => ws.mergeCells(1, c, 2, c));
    grid(ws, 1, 1, 2, cR, false); for (let c = 1; c <= cR; c++) ws.getCell(2, c).font = { bold: true };
    const L = c => ws.getColumn(c).letter;
    const block = (dateTxt, b) => {
      const rows = [["Designated", b.des], ["Undesignated", b.und], ["Total Designated + Undesignated", null], ["Contractor Reported", b.rep], ["% Deviation w.r.t Reported", null], ["Final after Reconciliation", b.fin]];
      if (opts.gps) rows.push(["GPS hauls (internal)", b.gps]);
      const r0 = ws.rowCount + 1;
      rows.forEach(([lab, v], k) => {
        const r = r0 + k, row = ws.getRow(r); row.getCell(5).value = lab;
        for (let i = 0; i < nS; i++) { const c = cS + i, col = L(c);
          if (lab.startsWith("Total")) row.getCell(c).value = { formula: `${col}${r0}+${col}${r0 + 1}`, result: b.tot[i] };
          else if (lab.startsWith("%")) row.getCell(c).value = { formula: `IFERROR(ROUND((${col}${r0 + 3}-${col}${r0 + 2})*100/${col}${r0 + 2},2),"-")`, result: b.dev[i] === null ? "-" : b.dev[i] };
          else row.getCell(c).value = v[i]; }
        if (lab.startsWith("Total") || lab.startsWith("Final")) row.eachCell({ includeEmpty: true }, x => { if (x.col >= 5 && x.col < cR) x.font = { bold: true }; });
        if (lab.startsWith("GPS")) for (let c = 5; c < cR; c++) ws.getCell(r, c).fill = YEL;
      });
      const r1 = r0 + rows.length - 1;
      ws.getCell(r0, 1).value = data.site; ws.getCell(r0, 2).value = data.nit; ws.getCell(r0, 3).value = data.contractor; ws.getCell(r0, 4).value = dateTxt;
      ws.getCell(r0, cR).value = S.reconRemark(b);
      [1, 2, 3, 4, cR].forEach(c => { ws.mergeCells(r0, c, r1, c); ws.getCell(r0, c).alignment = { vertical: "middle", horizontal: c === cR ? "left" : "center", wrapText: true }; });
      for (let r = r0; r <= r1; r++) for (let c = 1; c <= cR; c++) ws.getCell(r, c).border = BOX;
    };
    for (const [D, r] of res) { const bl = S.reconDay(D, r); bl.forEach(b => block(b.label === "Day" ? `${dmy(D.day)} — Day` : `${dmy(D.day)} / ${b.label}`, b)); }
    if (res.length > 1) block(`${dmy(res[0][0].day)} to ${dmy(res[res.length - 1][0].day)} — Total`, S.reconPeriod(res));
    add(ws, []);
    line(ws, "Designated / Undesignated: the VTS portal's trips of the NIT on the portal's own lead (Trip Management / NIT lead report), shift by trip start. Contractor Reported: each tipper-shift's trips in the slab of its reported lead. Slab: lower < lead ≤ upper. % Deviation w.r.t Reported = (Contractor Reported − Total) × 100 ÷ Total; \"-\" where the Total is 0. Final after Reconciliation: as on the SOP Daily Report — the hauls on the GPS track in the slab of the haul's lead (HQ formula at the actual loading and tip places; over 5 km in 4-5, the NIT's highest slab) plus trips given on justification for tipper-shifts with no GPS data (capped at the shift's reported − GPS hauls), in the slab chosen on the Justification Register.", { merge: cR, h: 60 });
    if (!(edits.sign && edits.sign.show === false)) { add(ws, []); add(ws, []); signatures(ws, ws.rowCount + 1, cR, reconSign(edits)); }
    [16, 9, 20, 14, 30].forEach((w, i) => (ws.getColumn(i + 1).width = w)); for (let i = 0; i < nS; i++) ws.getColumn(cS + i).width = 8; ws.getColumn(cR).width = 70;
    ws.views = [{ state: "frozen", ySplit: 2, xSplit: 5 }]; landscape(ws, false, "1:2");
    return wb;
  }
  const api = { buildWorkbook, buildReconWorkbook, DEFAULT_SIGN, reconSign, dmy, ist };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.SOPXlsx = api;
})(typeof self !== "undefined" ? self : this);
