/*
 * ORAI: CCB dump -> Fieldwork consumer master
 * Reads the monthly CCB file (DVVNL_DIV233422_MMM_YYYY .csv.gz or .csv)
 * as a stream (no need to unzip on the PC), keeps the chosen SDO(s),
 * and turns each consumer into the master columns the app reads:
 *  - dates 06-OCT-2026 -> 2026-10-06 (the app cannot read month names)
 *  - meter number without the leading '   (SERIAL_NBR, else METER_BADGE_NO)
 *  - closing reading from CLOSING_READING_KWH (CLOSE_READING is mostly blank)
 *  - substation cleaned to the six Orai town names; missing ones taken
 *    from other consumers of the same book
 *  - billing month = the month most consumers were last billed in
 * Result: rows (header + data) ready to be written as XLSX.
 */
(function (root) {
  const MASTER_COLS = ["ACCT_ID", "SDO_CODE", "NAME", "FATHER_NAME", "ADDRESS", "HABITAT_NAME", "SUPPLY_TYPE", "LOAD",
    "CON_STATUS", "SERIAL_NBR", "TOTAL_OUTSTANDING", "CLOSE_READING", "MOBILE_NO", "LAST_PAY_DATE", "LAST_PAY_AMOUNT",
    "SUBSTATION", "FEEDER", "LAST_BILL_DATE", "BILLING_MONTH", "REPORT_DATE", "CUR_ASSESSMENT", "PREV_ASSESSMENT",
    "THIS_MONTH_TOTAL_PAY_AMT", "MR_SOURCE_CD", "CONSUMPTION_CURR_MNTH", "CONSUMPTION_PREV_MNTH",
    "CONSUMPTION_PREV_TO_PREV_MNTH", "DOC", "BILL_BASIS", "METER_READ_REMARK", "SBM_MACHINE_ID", "BASE_ID"];
  const NUMERIC = new Set(["LOAD", "TOTAL_OUTSTANDING", "CLOSE_READING", "LAST_PAY_AMOUNT", "CUR_ASSESSMENT",
    "THIS_MONTH_TOTAL_PAY_AMT", "CONSUMPTION_CURR_MNTH", "CONSUMPTION_PREV_MNTH", "CONSUMPTION_PREV_TO_PREV_MNTH"]);
  const NEEDED = ["ACCT_ID", "SDO_CODE", "NAME", "FATHER_NAME", "ADDRESS", "HABITAT_NAME", "VILLAGE_NAME", "SUPPLY_TYPE",
    "LOAD", "CON_STATUS", "SERIAL_NBR", "METER_BADGE_NO", "TOTAL_OUTSTANDING", "CLOSING_READING_KWH", "CLOSE_READING",
    "LAST_OK_READING", "MOBILE_NO", "PAY_DATE", "PAY_AMT", "SUBSTATION", "FEEDER", "FEEDER_NAME", "LAST_BILL_DATE",
    "CURRENT_ASSESSMENT", "TOTAL_PAY_AMT", "MR_SOURCE_CD", "CONSUMPTION_CURR_MNTH", "CONSUMPTION_PREV_MNTH",
    "CONSUMPTION_PREV_TO_PREV_MNTH", "DOC", "BILL_BASIS", "METER_READ_REMARK", "SBM_MACHINE_ID", "BOOK_NO"];
  const DROP_STATUS = new Set(["NC REJECT", "REJECTED"]);
  const BILLABLE = new Set(["INSERVICE", "TD", "TD>6MONTH", "TD>6MONTHS"]);
  const MON = { JAN: 1, FEB: 2, MAR: 3, APR: 4, MAY: 5, JUN: 6, JUL: 7, AUG: 8, SEP: 9, OCT: 10, NOV: 11, DEC: 12 };
  const SUB_CODE = { "0132": "Kalpi Road", "0133": "Konch Naka", "0134": "Orai Kalpi Road", "0135": "Susheel Nagar",
    "0136": "Rahiya", "4273": "Karmer Road" };

  function ymd(v) {
    const m = /^\s*(\d{1,2})-([A-Za-z]{3})-(\d{2,4})/.exec(v || "");
    if (m) {
      const mo = MON[m[2].toUpperCase()];
      if (!mo) return "";
      let y = +m[3]; if (y < 100) y += 2000;
      return y + "-" + String(mo).padStart(2, "0") + "-" + String(+m[1]).padStart(2, "0");
    }
    return /^\d{4}-\d{2}-\d{2}/.test(v || "") ? v.slice(0, 10) : "";
  }

  function substation(ss, fd) {
    const s = String(ss || "").toUpperCase(), f = String(fd || "").toUpperCase();
    let m = /(?:^|[^0-9])2(013[2-6]|4273)(?![0-9])/.exec(s);
    if (m) return SUB_CODE[m[1]];
    for (const x of [s, f]) {
      m = /233422\d(013[2-6]|4273)/.exec(x);
      if (m) return SUB_CODE[m[1]];
    }
    const n = s.replace(/[^A-Z]/g, "");
    if (n.includes("ORAIKALPI")) return "Orai Kalpi Road";
    if (n.includes("KALPIROAD")) return "Kalpi Road";
    if (/K[OA]+N?CHNAKA/.test(n)) return "Konch Naka";
    if (/SU+S+H?[EIU]+L/.test(n)) return "Susheel Nagar";
    if (n.includes("RAHIYA")) return "Rahiya";
    if (n.includes("KARMER")) return "Karmer Road";
    /* Any other substation written as NAME_12345 (rural ones): keep its name. */
    const o = /^([A-Z][A-Z .\-]*[A-Z])_\d{5}$/.exec(s.trim());
    if (o) return o[1];
    return "";
  }

  const feederName = (fd, fname) => (String(fname || "").trim() || String(fd || "").replace(/_\d.*$/, "")).trim().toUpperCase();

  /* ---- streaming CSV reader (handles quotes, "" inside quotes, line breaks inside quotes) ---- */
  async function* csvRows(textStream) {
    const reader = textStream.getReader();
    let row = [], field = "", inQ = false, quoteSeen = false, carry = "";
    while (true) {
      const { value, done } = await reader.read();
      const chunk = done ? "" : carry + value;
      carry = "";
      const out = [];
      const n = chunk.length;
      let start = 0;
      for (let i = 0; i < n; i++) {
        const c = chunk.charCodeAt(i);
        if (inQ) {
          if (c === 34) {
            if (i + 1 >= n && !done) { field += chunk.slice(start, i); carry = '"'; start = n; break; }
            if (chunk.charCodeAt(i + 1) === 34) { field += chunk.slice(start, i + 1); i++; start = i + 1; }
            else { field += chunk.slice(start, i); inQ = false; start = i + 1; }
          }
        } else if (c === 34) {
          field += chunk.slice(start, i); inQ = true; quoteSeen = true; start = i + 1;
        } else if (c === 44) {
          row.push(field + chunk.slice(start, i)); field = ""; start = i + 1;
        } else if (c === 10 || c === 13) {
          row.push(field + chunk.slice(start, i)); field = "";
          if (c === 13 && chunk.charCodeAt(i + 1) === 10) i++;
          start = i + 1;
          if (row.length > 1 || row[0] !== "") out.push(row);
          row = [];
        }
      }
      if (start < n) field += chunk.slice(start);
      for (const r of out) yield r;
      if (done) {
        if (field !== "" || row.length) { row.push(field); yield row; }
        return;
      }
    }
    void quoteSeen;
  }

  function textStreamOf(file, onBytes) {
    let read = 0;
    const counted = file.stream().pipeThrough(new TransformStream({
      transform(chunk, ctl) { read += chunk.byteLength; onBytes && onBytes(read); ctl.enqueue(chunk); }
    }));
    const isGz = /\.gz$/i.test(file.name);
    if (isGz && typeof DecompressionStream === "undefined") throw new Error("This browser cannot open .gz files. Use Chrome, or unzip the file first.");
    return (isGz ? counted.pipeThrough(new DecompressionStream("gzip")) : counted).pipeThrough(new TextDecoderStream());
  }

  /* Is this a raw CCB dump (not an already-made master)? */
  async function isCcb(file) {
    try {
      const it = csvRows(textStreamOf(file));
      const { value } = await it.next();
      it.return && it.return();
      const h = (value || []).map(x => String(x).trim().toUpperCase());
      return h.includes("ACCT_ID") && h.includes("DIV_CODE") && h.includes("SERIAL_NBR") && h.includes("CURRENT_ASSESSMENT");
    } catch (e) { return false; }
  }

  /*
   * sdos: array of SDO codes to keep (e.g. ["SDO2334228"]); empty = all.
   * onProgress(fractionRead, rowsSeen)
   */
  async function convert(file, sdos, onProgress) {
    const keep = new Set((sdos || []).map(s => String(s).trim().toUpperCase()).filter(Boolean));
    let lastBytes = 0;
    const it = csvRows(textStreamOf(file, b => { lastBytes = b; }));
    const first = await it.next();
    const head = (first.value || []).map(x => String(x).trim().toUpperCase());
    const ix = {};
    for (const k of NEEDED) ix[k] = head.indexOf(k);
    for (const k of ["ACCT_ID", "SDO_CODE", "CON_STATUS"]) if (ix[k] < 0) throw new Error("Not a CCB file: column " + k + " missing.");
    const g = (r, k) => (ix[k] < 0 ? "" : String(r[ix[k]] ?? "").trim());

    const sdoCount = {}, seenAcct = new Set(), recs = [];
    let seen = 0, dropped = 0, dup = 0;
    const billMonth = {};
    let lastTick = 0;
    for await (const r of it) {
      seen++;
      const sdo = g(r, "SDO_CODE").toUpperCase();
      sdoCount[sdo] = (sdoCount[sdo] || 0) + 1;
      if (keep.size && !keep.has(sdo)) continue;
      const status = g(r, "CON_STATUS");
      if (DROP_STATUS.has(status.toUpperCase())) { dropped++; continue; }
      const acct = g(r, "ACCT_ID");
      if (!acct) continue;
      if (seenAcct.has(acct)) { dup++; continue; }
      seenAcct.add(acct);

      const lastBill = ymd(g(r, "LAST_BILL_DATE"));
      if (lastBill) billMonth[lastBill.slice(0, 7)] = (billMonth[lastBill.slice(0, 7)] || 0) + 1;
      const fdr = feederName(g(r, "FEEDER"), g(r, "FEEDER_NAME"));
      const hab = g(r, "HABITAT_NAME"), vil = g(r, "VILLAGE_NAME");
      recs.push({
        ACCT_ID: acct, SDO_CODE: sdo, NAME: g(r, "NAME"), FATHER_NAME: g(r, "FATHER_NAME"), ADDRESS: g(r, "ADDRESS"),
        HABITAT_NAME: hab ? hab.toUpperCase() : (vil.length > 2 ? vil.toUpperCase() : (fdr || "ORAI TOWN")),
        SUPPLY_TYPE: g(r, "SUPPLY_TYPE"), LOAD: g(r, "LOAD"), CON_STATUS: status,
        SERIAL_NBR: g(r, "SERIAL_NBR").replace(/^'+/, "").trim() || g(r, "METER_BADGE_NO"),
        TOTAL_OUTSTANDING: g(r, "TOTAL_OUTSTANDING"),
        CLOSE_READING: g(r, "CLOSING_READING_KWH") || g(r, "CLOSE_READING") || g(r, "LAST_OK_READING"),
        MOBILE_NO: g(r, "MOBILE_NO"), LAST_PAY_DATE: ymd(g(r, "PAY_DATE")), LAST_PAY_AMOUNT: g(r, "PAY_AMT"),
        SUBSTATION: substation(g(r, "SUBSTATION"), g(r, "FEEDER")), FEEDER: fdr, LAST_BILL_DATE: lastBill,
        CUR_ASSESSMENT: g(r, "CURRENT_ASSESSMENT"), PREV_ASSESSMENT: "", THIS_MONTH_TOTAL_PAY_AMT: g(r, "TOTAL_PAY_AMT"),
        MR_SOURCE_CD: g(r, "MR_SOURCE_CD"), CONSUMPTION_CURR_MNTH: g(r, "CONSUMPTION_CURR_MNTH"),
        CONSUMPTION_PREV_MNTH: g(r, "CONSUMPTION_PREV_MNTH"), CONSUMPTION_PREV_TO_PREV_MNTH: g(r, "CONSUMPTION_PREV_TO_PREV_MNTH"),
        DOC: ymd(g(r, "DOC")), BILL_BASIS: g(r, "BILL_BASIS"), METER_READ_REMARK: g(r, "METER_READ_REMARK"),
        SBM_MACHINE_ID: g(r, "SBM_MACHINE_ID"),
        _book: g(r, "BOOK_NO").replace(/'/g, "").trim(), _pay: ymd(g(r, "PAY_DATE"))
      });
      if (onProgress && seen - lastTick >= 5000) {
        lastTick = seen;
        onProgress(file.size ? lastBytes / file.size : 0, seen);
        await new Promise(res => setTimeout(res, 0));
      }
    }
    if (!recs.length) {
      throw new Error("No consumers found for " + (keep.size ? Array.from(keep).join(", ") : "this file") +
        ". SDO codes in the file: " + Object.keys(sdoCount).join(", "));
    }

    /* Missing substation: the substation most consumers of the same book have (60 % or more, 5+ consumers). */
    const book = {};
    for (const c of recs) if (c.SUBSTATION && c._book) {
      const b = book[c._book] || (book[c._book] = {});
      b[c.SUBSTATION] = (b[c.SUBSTATION] || 0) + 1;
    }
    const bookSub = {};
    for (const [b, counts] of Object.entries(book)) {
      const tot = Object.values(counts).reduce((a, x) => a + x, 0);
      const [best, n] = Object.entries(counts).sort((a, b2) => b2[1] - a[1])[0];
      if (tot >= 5 && n / tot >= 0.6) bookSub[b] = best;
    }
    let fromBook = 0;
    for (const c of recs) if (!c.SUBSTATION) {
      if (bookSub[c._book]) { c.SUBSTATION = bookSub[c._book]; fromBook++; } else c.SUBSTATION = "Not mapped";
    }

    /* Billing month = month in which most consumers were last billed; report date = last bill / payment date in it. */
    const billingMonth = Object.entries(billMonth).sort((a, b) => b[1] - a[1])[0]?.[0] || "";
    let reportDate = "";
    for (const c of recs) for (const d of [c.LAST_BILL_DATE, c._pay]) if (d && d.slice(0, 7) === billingMonth && d > reportDate) reportDate = d;

    const stats = { seen, kept: recs.length, dropped, dup, fromBook, sdoCount, billingMonth, reportDate,
      billable: 0, billed: 0, collection: 0, bySub: {} };
    const rows = [MASTER_COLS];
    /* One id per master: daily billed/unbilled updates are built for, and only applied to, this master. */
    const baseId = "ORAI-" + (billingMonth || "M") + "-" + Date.now().toString(36);
    stats.baseId = baseId;
    for (const c of recs) {
      c.BILLING_MONTH = billingMonth; c.REPORT_DATE = reportDate; c.BASE_ID = baseId;
      if (BILLABLE.has(c.CON_STATUS.toUpperCase().replace(/\s+/g, ""))) {
        stats.billable++;
        if (c.LAST_BILL_DATE.slice(0, 7) === billingMonth) stats.billed++;
      }
      stats.collection += Number(c.THIS_MONTH_TOTAL_PAY_AMT) || 0;
      stats.bySub[c.SUBSTATION] = (stats.bySub[c.SUBSTATION] || 0) + 1;
      rows.push(MASTER_COLS.map(k => {
        const v = c[k];
        if (NUMERIC.has(k) && v !== "" && isFinite(Number(v))) return Number(v);
        return v;
      }));
    }
    if (onProgress) onProgress(1, seen);
    return { rows, stats };
  }

  /* ---- XLSX is a ZIP. SheetJS compresses weakly (a 72k-consumer master ~33 MB); re-pack it with the
     browser's own deflate (~12 MB) so the upload and every phone's download are much smaller. ---- */
  const CRC_T = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(u8) { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = CRC_T[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  async function deflateRaw(u8) {
    const s = new Blob([u8]).stream().pipeThrough(new CompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(s).arrayBuffer());
  }
  async function inflateRaw(u8) {
    const s = new Blob([u8]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(s).arrayBuffer());
  }
  async function repackZip(zip) {
    if (typeof CompressionStream === "undefined") return zip;
    const dv = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
    let eocd = -1;
    for (let i = zip.length - 22; i >= 0; i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) return zip;
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const files = [];
    for (let k = 0; k < count; k++) {
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = zip.subarray(p + 46, p + 46 + nlen);
      const dstart = lho + 30 + dv.getUint16(lho + 26, true) + dv.getUint16(lho + 28, true);
      let data = zip.subarray(dstart, dstart + csize);
      if (method === 8) data = await inflateRaw(data); else if (method !== 0) return zip;
      files.push({ name, data });
      p += 46 + nlen + xlen + clen;
    }
    return zipFiles(files);
  }

  async function zipFiles(files) {
    const parts = [], central = [];
    let off = 0;
    for (const f of files) {
      const comp = await deflateRaw(f.data), crc = crc32(f.data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(8, 8, true);
      h.setUint16(12, 0x21, true);                       /* 1980-01-01 */
      h.setUint32(14, crc, true); h.setUint32(18, comp.length, true); h.setUint32(22, f.data.length, true);
      h.setUint16(26, f.name.length, true);
      parts.push(new Uint8Array(h.buffer), f.name, comp);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(10, 8, true);
      c.setUint16(14, 0x21, true); c.setUint32(16, crc, true); c.setUint32(20, comp.length, true);
      c.setUint32(24, f.data.length, true); c.setUint16(28, f.name.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), f.name);
      off += 30 + f.name.length + comp.length;
    }
    const csz = central.reduce((a, x) => a + x.length, 0);
    const e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, csz, true); e.setUint32(16, off, true);
    return new Uint8Array(await new Blob([...parts, ...central, new Uint8Array(e.buffer)]).arrayBuffer());
  }
  /* ---- Write the master as a plain one-sheet XLSX directly (much faster than a general Excel library). ---- */
  const esc = v => String(v).replace(/[&<>"]/g, ch => ch === "&" ? "&amp;" : ch === "<" ? "&lt;" : ch === ">" ? "&gt;" : "&quot;")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  function colName(i) { let s = ""; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - m - 1) / 26; } return s; }
  async function buildXlsx(rows) {
    const enc = new TextEncoder();
    const cols = rows[0].map((_, i) => colName(i));
    const chunks = [];
    let buf = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r], rn = r + 1;
      let x = '<row r="' + rn + '">';
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        if (v === "" || v == null) continue;
        const ref = cols[c] + rn;
        x += typeof v === "number" && isFinite(v)
          ? '<c r="' + ref + '"><v>' + v + '</v></c>'
          : '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
      }
      buf += x + '</row>';
      if (buf.length > 1 << 20) { chunks.push(enc.encode(buf)); buf = ""; }
    }
    chunks.push(enc.encode(buf + '</sheetData></worksheet>'));
    let total = 0; for (const c of chunks) total += c.length;
    const sheet = new Uint8Array(total); let o = 0; for (const c of chunks) { sheet.set(c, o); o += c.length; }
    const parts = [
      ["[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'],
      ["_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
      ["xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Master" sheetId="1" r:id="rId1"/></sheets></workbook>'],
      ["xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'],
      ["xl/worksheets/sheet1.xml", sheet]
    ];
    return zipFiles(parts.map(([n, d]) => ({ name: enc.encode(n), data: typeof d === "string" ? enc.encode(d) : d })));
  }


  root.OraiCCB = { isCcb, convert, repackZip, buildXlsx, MASTER_COLS };
})(typeof window !== "undefined" ? window : globalThis);
