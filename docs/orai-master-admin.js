/*
 * ORAI: Admin > Master Data
 * Upload a new consumer master (XLSX, or CSV converted to XLSX here),
 * check its columns first, choose who gets it (All / a subdivision),
 * and see / switch the masters already on the server.
 * Phones load a new master by themselves on their next update check.
 */
(() => {
  const CHUNK = 3 * 1000 * 1000;           /* raw bytes per part; a multiple of 3 so Base64 parts join cleanly */

  const norm = h => String(h ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");

  /* Same column list the app checks when it imports a master. */
  const REQUIRED = [
    ["ACCT_ID", ["acct_id"]],
    ["SDO_CODE", ["sdo_code"]],
    ["NAME", ["name"]],
    ["FATHER_NAME", ["father_name"]],
    ["ADDRESS", ["address"]],
    ["Village (FEEDED VILLAGE NAME / HABITAT_NAME)", ["feeded village name", "habitat_name"]],
    ["SUPPLY_TYPE", ["supply_type"]],
    ["LOAD", ["load"]],
    ["CONNECTION STATUS / CON_STATUS", ["connection status", "con_status"]],
    ["METER NO / SERIAL_NBR", ["meter no", "serial_nbr"]],
    ["TOTAL_OUTSTANDING", ["total outstanding", "total_outstanding"]],
    ["CURRENT READING / CLOSE_READING", ["current reading", "close_reading"]],
    ["MOBILE_NO", ["mobile number", "mobile_no"]],
    ["LAST_PAY_DATE", ["last_pay_date", "updated_last_pay_date"]],
    ["LAST_PAY_AMOUNT", ["last_pay_amount", "updated_last_pay_amount"]]
  ];
  /* Needed for Progress > Billed / Unbilled (and collection) figures. */
  const BILLING = [
    ["LAST_BILL_DATE", ["last_bill_date"], "billed / unbilled (essential)"],
    ["BILLING_MONTH", ["billing_month"], "which month counts as billed (else the current month)"],
    ["REPORT_DATE", ["report_date"], "\"data up to\" date (else today)"],
    ["SUBSTATION", ["substation"], "substation-wise rows"],
    ["FEEDER", ["feeder"], "feeder-wise rows"],
    ["CUR_ASSESSMENT", ["cur_assessment"], "this month's assessment"],
    ["PREV_ASSESSMENT", ["prev_assessment"], "last month's assessment"],
    ["THIS_MONTH_TOTAL_PAY_AMT", ["this_month_total_pay_amt"], "collection this month"],
    ["MR_SOURCE_CD", ["mr_source_cd"], "smart / MR source split"]
  ];

  const $ = id => document.getElementById(id);
  let picked = null;                       /* {bytes, name, consumers, billing_month, report_date} */

  function html() {
    const card = $("adminCard");
    if (!card || $("omSection")) return;
    const sec = document.createElement("div");
    sec.id = "omSection";
    sec.innerHTML = `
      <style>
        #omSection{margin:22px 0;border-top:1px solid #e2e8f0;padding-top:12px}
        #omSection .om-box{border:1px solid #e2e8f0;border-radius:10px;padding:10px;margin:8px 0;font-size:13px}
        #omSection .om-ok{color:#087443;font-weight:700}
        #omSection .om-bad{color:#b42318;font-weight:700}
        #omSection .om-warn{color:#b45309;font-weight:700}
        #omSection ul{margin:6px 0 0 18px;padding:0}
        #omSection table{border-collapse:collapse;width:100%;font-size:12.5px}
        #omSection th,#omSection td{border-bottom:1px solid #edf2f7;padding:6px;text-align:left;vertical-align:top}
        #omSection th{background:#f0fdfa;color:#0f766e}
        #omSection .om-bar{height:8px;background:#e2e8f0;border-radius:6px;overflow:hidden;margin-top:8px}
        #omSection .om-bar div{height:100%;width:0;background:#0d9488;transition:width .2s}
      </style>
      <h2>Master Data(मास्टर डेटा)</h2>
      <p class="muted" style="font-size:13px">Choose the monthly CCB file (DVVNL_DIV233422_MMM_YYYY .csv.gz) as it is — no need to
        unzip or edit it — or a ready master XLSX. Every phone in the chosen
        scope loads it by itself on its next update check. (नया मास्टर अपलोड करें)</p>

      <div id="omServer" class="om-box">Checking the Orai server…</div>
      <label for="omSdo">SDOs to include (only for the raw CCB file) — ALL = full master</label>
      <input id="omSdo" value="ALL" placeholder="ALL = full master (every SDO) · or e.g. SDO2334228 · several: SDO2334228, SDO2334226">

      <label for="omFile">Master file — CCB dump (.csv.gz / .csv) or ready master (.xlsx)</label>
      <input id="omFile" type="file" accept=".gz,.csv,.xlsx,.xls">
      <div id="omCheck"></div>

      <label for="omScope">Who gets this master</label>
      <select id="omScope"><option value="ALL">All users (whole Orai)</option></select>

      <label for="omNotes">Note (optional)</label>
      <input id="omNotes" placeholder="e.g. November 2026 master">

      <button id="omUploadBtn" class="primary" type="button" disabled>UPLOAD MASTER DATA</button>
      <div class="om-bar hidden" id="omBarWrap"><div id="omBar"></div></div>
      <div id="omStatus" class="status"></div>

      <h3>Masters on the server</h3>
      <button id="omRefreshBtn" class="secondary" type="button">REFRESH LIST</button>
      <div id="omList" class="om-box">Not loaded yet.</div>

      <h3>Billed / Unbilled report</h3>
      <p class="muted" style="font-size:13px">Open <b>PROGRESS</b> on the home screen → <b>Billing</b> section:
        billable, billed, billing % and unbilled (SDO / substation / feeder / village, by category and load);
        tap any number for the consumer list. It needs the LAST_BILL_DATE column in the master.</p>
      <button id="omProgressBtn" class="secondary" type="button">OPEN BILLED / UNBILLED REPORT</button>
    `;
    const usersHead = Array.from(card.querySelectorAll("h3")).find(h => /Users/.test(h.textContent));
    card.insertBefore(sec, usersHead || null);

    $("omFile").addEventListener("change", onPick);
    $("omSdo").addEventListener("change", () => { if ($("omFile").files[0]) onPick(); });
    $("omUploadBtn").addEventListener("click", upload);
    $("omRefreshBtn").addEventListener("click", () => { checkServer(); loadList(); });
    $("omProgressBtn").addEventListener("click", () => {
      const b = $("progressBtn");
      if (b) b.click();
    });
  }

  async function onPick() {
    picked = null;
    $("omUploadBtn").disabled = true;
    const box = $("omCheck");
    const f = $("omFile").files[0];
    if (!f) { box.innerHTML = ""; return; }
    if (typeof XLSX === "undefined") {
      box.innerHTML = `<div class="om-box om-bad">The Excel reader did not load. Check internet and reopen the app.</div>`;
      return;
    }
    box.innerHTML = `<div class="om-box">Checking ${escapeHtml(f.name)} (${(f.size / 1048576).toFixed(1)} MB)…</div>`;
    await new Promise(r => setTimeout(r, 50));
    try {
      if (window.OraiCCB && (/\.gz$/i.test(f.name) || (/\.csv$/i.test(f.name) && await OraiCCB.isCcb(f)))) {
        return await fromCcb(f, box);
      }
      const raw = new Uint8Array(await f.arrayBuffer());
      const isCsv = /\.csv$/i.test(f.name);
      /* CSV: keep every value as text so account / meter / mobile numbers stay exact. */
      const wb = XLSX.read(raw, isCsv ? { type: "array", raw: true, dense: true }
                                      : { type: "array", dense: true, cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      if (!ws || !ws["!ref"]) throw new Error("The file has no data sheet.");
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true, range: 0 });
      await finish(rows, f.name.replace(/\.(csv|xls)$/i, ".xlsx"), /\.xlsx$/i.test(f.name) ? raw : null, wb, "");
    } catch (e) {
      box.innerHTML = `<div class="om-box om-bad">Could not read the file: ${escapeHtml(e.message || String(e))}</div>`;
    }
  }

  /* Raw CCB dump -> master rows, converted here on this device. */
  async function fromCcb(f, box) {
    const txt = $("omSdo").value.trim().toUpperCase();
    const sdos = (!txt || txt === "ALL") ? [] : txt.split(/[\s,;]+/).filter(Boolean);
    const t0 = Date.now();
    const show = (p, n) => {
      box.innerHTML = `<div class="om-box">Converting CCB file ${escapeHtml(f.name)}… ${Math.round(100 * p)}% ` +
        `(${n.toLocaleString("en-IN")} rows read) — keep this screen open.<div class="om-bar"><div style="width:${Math.round(100 * p)}%"></div></div></div>`;
    };
    show(0, 0);
    const { rows, stats } = await OraiCCB.convert(f, sdos, show);
    box.innerHTML = `<div class="om-box">Making the master file (${stats.kept.toLocaleString("en-IN")} consumers)…</div>`;
    await new Promise(r => setTimeout(r, 30));
    const N = n => Number(n).toLocaleString("en-IN");
    const subs = Object.entries(stats.bySub).sort((a, b) => b[1] - a[1])
      .map(([k, v]) => `${escapeHtml(k)} ${N(v)}`).join(" · ");
    const sdoTxt = Object.entries(stats.sdoCount).map(([k, v]) => `${escapeHtml(k)} ${N(v)}`).join(" · ");
    const info = `<div class="om-box"><b>CCB file converted</b> in ${Math.round((Date.now() - t0) / 1000)} s<br>` +
      `SDOs in file: ${sdoTxt}<br>Kept: <b>${sdos.length ? escapeHtml(sdos.join(", ")) : "all SDOs"}</b> — ` +
      `${N(stats.kept)} consumers (${N(stats.dropped)} NC Reject / Rejected left out)<br>` +
      `Billing month <b>${escapeHtml(stats.billingMonth)}</b>, data up to <b>${escapeHtml(stats.reportDate)}</b><br>` +
      `Billable ${N(stats.billable)} · Billed ${N(stats.billed)} · <b>Unbilled ${N(stats.billable - stats.billed)}</b> · ` +
      `Collection ₹${N(Math.round(stats.collection))}<br>` +
      `Substations: ${subs}${stats.fromBook ? ` <small>(${N(stats.fromBook)} placed by book number)</small>` : ""}</div>`;
    const name = "Orai_Master_" + (sdos.length === 1 ? sdos[0] + "_" : "") + (stats.billingMonth || "CCB") + ".xlsx";
    await finish(rows, name, null, null, info);
  }

  /* Column check, then keep the XLSX bytes ready for upload. */
  async function finish(rows, outName, rawXlsx, wb, extraHtml) {
    const box = $("omCheck");
    const header = (rows[0] || []).map(norm);
    const has = names => names.some(n => header.includes(norm(n)));
    const missing = REQUIRED.filter(([, n]) => !has(n)).map(([l]) => l);
    const billOk = BILLING.filter(([, n]) => has(n));
    const billMissing = BILLING.filter(([, n]) => !has(n));
    const data = rows.slice(1).filter(r => r.some(v => String(v ?? "").trim() !== ""));
    const col = name => header.indexOf(norm(name));
    const first = name => {
      const i = col(name);
      if (i < 0) return "";
      const v = data.find(r => String(r[i] ?? "").trim())?.[i];
      if (v instanceof Date && !isNaN(v)) {                 /* Excel date cell -> 2026-10-01 */
        const d = new Date(v.getTime() - v.getTimezoneOffset() * 60000);
        return d.toISOString().slice(0, 10);
      }
      return String(v ?? "").trim();
    };
    const acctCol = col("acct_id");
    const accounts = acctCol < 0 ? 0 : data.filter(r => String(r[acctCol] ?? "").trim()).length;

    let out = extraHtml + `<div class="om-box"><b>${escapeHtml(outName)}</b> — ${accounts.toLocaleString("en-IN")} consumers<br>`;
    if (missing.length) {
      out += `<span class="om-bad">Cannot be used: required columns missing</span><ul>` +
        missing.map(m => `<li>${escapeHtml(m)}</li>`).join("") + `</ul>`;
    } else {
      out += `<span class="om-ok">✓ All required columns present — search and survey will work.</span>`;
    }
    if (!has(["last_bill_date"])) {
      out += `<br><span class="om-warn">⚠ No LAST_BILL_DATE column — the Billed / Unbilled report will show 0 billed.</span>`;
    } else {
      out += `<br><span class="om-ok">✓ Billed / Unbilled report: ${billOk.length} of ${BILLING.length} report columns present.</span>`;
    }
    if (billMissing.length && !missing.length) {
      out += `<ul>` + billMissing.map(([l, , w]) => `<li>Missing ${escapeHtml(l)} — ${escapeHtml(w)}</li>`).join("") + `</ul>`;
    }
    let bm = first("billing_month");
    const rd = first("report_date");
    if (/^\d{4}-\d{2}-\d{2}$/.test(bm)) bm = bm.slice(0, 7);
    if (bm || rd) out += `<br>Billing month: <b>${escapeHtml(bm || "-")}</b> · Data up to: <b>${escapeHtml(rd || "-")}</b>`;
    out += `</div>`;
    box.innerHTML = out;
    if (missing.length) return;
    if (!accounts) { box.innerHTML += `<div class="om-box om-bad">No consumer rows found.</div>`; return; }

    /* Phones read XLSX: anything else is written as XLSX here. */
    let bytes = rawXlsx;
    if (!bytes && !wb && window.OraiCCB && typeof CompressionStream !== "undefined") {
      box.innerHTML += `<div class="om-box" id="omMaking">Making the XLSX file…</div>`;
      await new Promise(r => setTimeout(r, 30));
      bytes = await OraiCCB.buildXlsx(rows);
      const mk = $("omMaking"); if (mk) mk.remove();
    }
    if (!bytes) {
      if (!wb) {
        wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Master");
      }
      bytes = new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx", compression: false, bookSST: true }));
      if (window.OraiCCB) bytes = await OraiCCB.repackZip(bytes);
    }
    picked = { bytes, name: outName, consumers: accounts, billing_month: bm, report_date: rd };
    box.innerHTML += `<div class="om-box">Ready to upload: ${(bytes.length / 1048576).toFixed(1)} MB.
      <button class="small-btn" type="button" id="omSaveBtn">SAVE COPY OF MASTER</button></div>`;
    $("omSaveBtn").addEventListener("click", () => {
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      a.download = outName; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    });
    $("omUploadBtn").disabled = false;
  }

  function toB64(u8) {
    let s = "";
    for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
    return btoa(s);
  }

  const OLD_SERVER_HELP = "The Orai server is still the OLD version. In Apps Script: (1) select all code and paste the new " +
    "orai-server.gs, Save; (2) run setupOrai and allow Google Drive access; (3) Deploy > Manage deployments > pencil > " +
    "Version: New version > Deploy. Then reopen this screen.";
  const DRIVE_HELP = "The server is not allowed to use Google Drive yet. In Apps Script run setupOrai once and allow the " +
    "permissions, then Deploy > Manage deployments > pencil > Version: New version > Deploy.";

  /* Turns server / network errors into what the admin has to do. */
  function explain(msg) {
    const m = String(msg || "");
    if (m.startsWith("\u2063")) return m;                       /* already explained */
    const out = (() => {
      if (/not set up on the Orai server/i.test(m)) return OLD_SERVER_HELP;
      if (/authori[sz]ation|permission|DriveApp|Drive access|scopes/i.test(m)) return DRIVE_HELP + " (" + m + ")";
      if (/Unexpected token|JSON|<!DOCTYPE|<html/i.test(m)) return "The server sent back an error page instead of data — usually " +
        "Drive permission not given or the deployment not updated. " + DRIVE_HELP;
      if (/exceeded maximum execution time|timed out/i.test(m)) return "The server took too long. Please try again.";
      if (/Failed to fetch|Network problem|NetworkError/i.test(m)) return "Could not reach the server — check internet. (" + m.replace(/^Network problem: /, "") + ")";
      if (/Login invalid|Admin access required/i.test(m)) return "Your login has expired. Log out and log in again as admin.";
      return m;
    })();
    return "\u2063" + out;
  }

  async function call(action, extra, tries = 3) {
    let last;
    for (let t = 1; t <= tries; t++) {
      try {
        const r = await adminRequest(action, extra);
        if (r && (r.success || r.code === "AUTH" || t === tries ||
            /not set up on the Orai server|Drive access|authori/i.test(r.message || ""))) {
          if (r && !r.success) r.message = explain(r.message);
          return r;
        }
        last = r;
      } catch (e) { last = { success: false, message: explain("Network problem: " + (e.message || e)) }; }
      if (t < tries) await new Promise(r => setTimeout(r, 2000 * t));
    }
    if (last && !last.success) last.message = explain(last.message);
    return last;
  }

  async function checkServer() {
    const box = $("omServer");
    if (!box) return false;
    box.innerHTML = "Checking the Orai server…";
    const r = await call("server_info", {}, 1);
    if (r && r.success) {
      box.innerHTML = `<span class="om-ok">✓ Orai server ready for Master Data</span> <small>(${escapeHtml(r.version || "")})</small>`;
      return true;
    }
    box.dataset.msg = ((r && r.message) || "No answer from server.").replace(/^\u2063/, "");
    box.innerHTML = `<span class="om-bad">✗ Master Data upload will not work yet.</span><br>${escapeHtml(box.dataset.msg)}`;
    return false;
  }

  async function upload() {
    if (!picked || !state.isAdmin) return;
    const scope = $("omScope").value || "ALL";
    if (!confirm(`Upload ${picked.name} (${picked.consumers.toLocaleString("en-IN")} consumers) for ` +
        `${scope === "ALL" ? "ALL users" : scope}?\nIt replaces the current master for that scope.`)) return;

    const btn = $("omUploadBtn"), bar = $("omBar");
    btn.disabled = true;
    $("omBarWrap").classList.remove("hidden");
    bar.style.width = "0%";
    const say = (t, k) => setStatus("omStatus", t, k);
    try {
      say("Checking the server…");
      if (!(await checkServer())) throw new Error("\u2063" + ($("omServer").dataset.msg || "Server not ready."));
      say("Starting upload…");
      const st = await call("master_upload_start", {});
      if (!st.success) throw new Error(st.message || "Could not start the upload.");
      const total = picked.bytes.length, parts = Math.ceil(total / CHUNK);
      for (let i = 0; i < parts; i++) {
        say(`Uploading part ${i + 1} of ${parts}… keep this screen open.`);
        const r = await call("master_upload_chunk", {
          upload_id: st.upload_id, index: i,
          data: toB64(picked.bytes.subarray(i * CHUNK, Math.min(total, (i + 1) * CHUNK)))
        });
        if (!r.success) throw new Error(r.message || `Part ${i + 1} failed.`);
        bar.style.width = Math.round(90 * (i + 1) / parts) + "%";
      }
      say("Saving on the server…");
      const fin = await call("master_upload_finish", {
        upload_id: st.upload_id, parts, total_size: total, file_name: picked.name, scope,
        consumers: picked.consumers, billing_month: picked.billing_month, report_date: picked.report_date,
        notes: $("omNotes").value.trim()
      }, 1);
      if (!fin.success) throw new Error(fin.message || "Upload could not be saved.");
      bar.style.width = "100%";
      say("✓ " + fin.message, "ok");
      $("omFile").value = ""; $("omCheck").innerHTML = ""; $("omNotes").value = ""; picked = null;
      loadList();
    } catch (e) {
      say(explain(e.message || String(e)).replace(/^\u2063/, ""), "error");
      btn.disabled = !picked;
    }
  }

  async function loadList() {
    const box = $("omList");
    if (!box || !state.isAdmin) return;
    box.textContent = "Loading…";
    const r = await call("master_list", {}, 2);
    if (!r || !r.success) { box.textContent = String((r && r.message) || "Could not load the list.").replace(/^\u2063/, ""); return; }

    const sel = $("omScope"), cur = sel.value;
    sel.innerHTML = (r.scopes || ["ALL"]).map(s =>
      `<option value="${escapeHtml(s)}">${s === "ALL" ? "All users (whole Orai)" : "Only " + escapeHtml(s)}</option>`).join("");
    if (cur) sel.value = cur;

    const list = r.masters || [];
    if (!list.length) { box.innerHTML = `<span class="om-warn">No master uploaded yet. Phones cannot search consumers until one is uploaded.</span>`; return; }
    const d = v => { const x = new Date(v); return isNaN(x) ? escapeHtml(v || "") : x.toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" }); };
    box.innerHTML = `<div style="overflow-x:auto"><table><thead><tr><th>Status</th><th>For</th><th>File</th>` +
      `<th>Consumers</th><th>Billing month</th><th>Uploaded</th><th></th></tr></thead><tbody>` +
      list.map(m => `<tr><td>${m.active ? '<span class="om-ok">ACTIVE</span>' : "old"}</td>` +
        `<td>${escapeHtml(m.scope === "ALL" ? "All" : m.scope)}</td>` +
        `<td>${escapeHtml(m.file_name)}<br><small>${(m.file_size / 1048576).toFixed(1)} MB${m.notes ? " · " + escapeHtml(m.notes) : ""}</small></td>` +
        `<td>${m.consumers ? Number(m.consumers).toLocaleString("en-IN") : "-"}</td>` +
        `<td>${escapeHtml(m.billing_month || "-")}</td>` +
        `<td>${d(m.last_updated)}<br><small>${escapeHtml(m.uploaded_by)}</small></td>` +
        `<td>${m.active ? "" : `<button class="small-btn" type="button" data-om-act="${escapeHtml(m.master_file_id)}">MAKE ACTIVE</button>`}</td></tr>`
      ).join("") + `</tbody></table></div>`;
    box.querySelectorAll("[data-om-act]").forEach(b => b.addEventListener("click", async () => {
      if (!confirm("Make this older file the active master again? Phones will switch back to it.")) return;
      const x = await call("master_set_active", { master_file_id: b.getAttribute("data-om-act") }, 1);
      alert(String(x.message || (x.success ? "Done." : "Failed.")).replace(/^\u2063/, ""));
      loadList();
    }));
  }

  function init() {
    html();
    const adminBtn = $("adminBtn");
    if (adminBtn) adminBtn.addEventListener("click", () => setTimeout(() => { checkServer(); loadList(); }, 300));
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
