/* Statement import: reads a Canara Bank e-passbook PDF in the browser.
 * Nothing leaves the device except the transactions you confirm, which go to your own Google Sheet.
 * Works in the browser (window.Statement) and in Node for tests (module.exports).
 */
(function (root) {
  'use strict';

  /** Turn a PDF into text lines, page by page, in reading order. */
  async function pdfToLines(pdfjsLib, data, password) {
    const doc = await pdfjsLib.getDocument({ data, password: password || undefined }).promise;
    const lines = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const tc = await page.getTextContent();
      const rows = [];
      for (const it of tc.items) {
        if (!it.str || !it.str.trim()) continue;
        const x = it.transform[4], y = it.transform[5];
        let row = rows.find(r => Math.abs(r.y - y) < 2.5);
        if (!row) { row = { y, items: [] }; rows.push(row); }
        row.items.push({ x, s: it.str, w: it.width || 0 });
      }
      rows.sort((a, b) => b.y - a.y);
      for (const r of rows) {
        r.items.sort((a, b) => a.x - b.x);
        let s = '', end = null;
        for (const it of r.items) {
          if (end !== null) s += (it.x - end > 3 ? '   ' : '');
          s += it.s; end = it.x + it.w;
        }
        lines.push(s.trim());
      }
      lines.push('<<page>>');
    }
    return lines;
  }

  const num = s => parseFloat(String(s).replace(/,/g, ''));
  const DATE_LINE = /^(\d{2}-\d{2}-\d{4})\s*(.*?)\s+([\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})\s*$/;
  const SKIP = /^(Date\s+Particulars|page\s*\d+|<<page>>|Statement for|Branch Code|Customer Id|Branch Name|Name\s|Phone|Address|Product|IFSC|THATAMPARAI|TAMIL NADU|Road,)/i;

  /** Parse Canara e-passbook lines into transactions. Deposit vs withdrawal is decided from the running balance. */
  function parseCanara(lines) {
    let opening = null, prev = null, buf = [], cur = null;
    const rows = [], warnings = [];
    let account = null, period = null;
    for (const raw of lines) {
      const l = raw.trim();
      if (!l) continue;
      const m0 = l.match(/Statement for A\/c\s+(\S+)\s+between\s+(\S+)\s+and\s+(\S+)/i);
      if (m0) { account = m0[1]; period = [m0[2], m0[3]]; continue; }
      if (/Opening Balance/i.test(l)) { opening = prev = num(l.match(/(-?[\d,]+\.\d{2})\s*$/)[1]); buf = []; continue; }
      if (/END OF STATEMENT/i.test(l)) break;
      if (SKIP.test(l)) continue;
      const m = l.match(DATE_LINE);
      if (m && prev !== null) {
        if (m[2]) buf.push(m[2]);
        cur = { date: m[1], amount: num(m[3]), balance: num(m[4]) };
        continue;
      }
      if (/^Chq:/i.test(l)) {
        if (cur) {
          cur.narration = buf.map(x => x.trim()).join('');
          const up = Math.round((cur.balance - prev) * 100), amt = Math.round(cur.amount * 100);
          if (up === amt) cur.type = 'in';
          else if (-up === amt) cur.type = 'out';
          else { cur.type = /UPI\/CR|NEFT CR|CASH DEPOSIT|SBINT|REV:/i.test(cur.narration) ? 'in' : 'out'; warnings.push('Balance did not reconcile on ' + cur.date); }
          prev = cur.balance; rows.push(cur);
        }
        cur = null; buf = []; continue;
      }
      if (prev !== null) buf.push(l);
    }
    return { account, period, opening, closing: prev, rows, warnings };
  }

  function isoDate(d) { const [dd, mm, yy] = d.split('-'); return yy + '-' + mm + '-' + dd; }

  function counterparty(n) {
    if (/^UPI\//.test(n)) return (n.split('/')[3] || '').trim();
    if (/ATM CASH/.test(n)) { const p = n.split('-'); return ('ATM ' + (p[2] || '').replace(/[^A-Z ]/g, '')).slice(0, 32).trim(); }
    if (/IGRAPH/.test(n)) return 'IGRAPH TECHNOLOGIES';
    if (/SREESARADHAMBAL/.test(n)) return 'SREE SARADHAMBAL AUTOMOBILES';
    if (/SUNDARAM/.test(n)) return 'SUNDARAM FINANCE';
    if (/SLPEINFOT/.test(n)) return 'SLPE / IDFC PAYOUT';
    return n.replace(/[\d\/:]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 32);
  }

  /** Default category when the payee has never been categorised before. */
  function ruleCategory(type, n) {
    if (type === 'in') {
      if (/IGRAPH/.test(n) || (/RAJAVEL/.test(n) && /SALARY/.test(n))) return 'Salary';
      if (/CASH DEPOSIT/.test(n)) return 'Cash deposit';
      if (/SBINT/.test(n)) return 'Interest';
      return 'Uncategorised';
    }
    if (/SREESARA|SUNDARAM/.test(n)) return 'Vehicle & loan';
    if (/ATM CASH/.test(n)) return 'ATM cash';
    if (/GOOGLEPL/.test(n)) return 'Subscriptions';
    if (/BHARTIAI|AIRTEL|\/JIO\//.test(n)) return 'Mobile recharge';
    if (/STATETRA/.test(n)) return 'Bus travel';
    if (/CHARGES|SMS|CASH HANDLING|RTGS 00|ECS MANDATE|CIBIL|TRUZTEE/.test(n)) return 'Bank charges';
    return 'Uncategorised';
  }

  function hash(s) { // FNV-1a, 12 hex chars; stable id for the same bank line
    let h1 = 0x811c9dc5, h2 = 0x01000193;
    for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); h1 = Math.imul(h1 ^ c, 16777619); h2 = Math.imul(h2 ^ c, 2246822519); }
    return ((h1 >>> 0).toString(16).padStart(8, '0') + (h2 >>> 0).toString(16).padStart(8, '0')).slice(0, 12);
  }

  /**
   * Build transactions ready to import.
   * existing: transactions already in the sheet. Used to (1) learn payee → category and
   * (2) skip rows already present: rows are matched on date + type + amount, counting duplicates,
   * so a statement that overlaps an earlier one never adds the same payment twice.
   */
  function prepare(parsed, existing) {
    const learned = {};
    existing.slice().sort((a, b) => (a.date < b.date ? -1 : 1)).forEach(t => {
      if (t.counterparty && t.category && t.category !== 'Uncategorised') learned[t.counterparty.toUpperCase()] = t.category;
    });
    const have = {};
    existing.forEach(t => { const k = t.date + '|' + t.type + '|' + Number(t.amount).toFixed(2); have[k] = (have[k] || 0) + 1; });
    const out = [], dupes = [];
    parsed.rows.forEach(r => {
      const date = isoDate(r.date), cp = counterparty(r.narration);
      const k = date + '|' + r.type + '|' + r.amount.toFixed(2);
      if (have[k] > 0) { have[k]--; dupes.push(r); return; }
      const cat = learned[cp.toUpperCase()] || ruleCategory(r.type, r.narration);
      out.push({ id: 'cnrb-' + hash(date + '|' + r.type + '|' + r.amount.toFixed(2) + '|' + r.balance.toFixed(2) + '|' + r.narration.slice(0, 40)),
        date, type: r.type, category: cat, counterparty: cp, amount: r.amount, note: '', source: 'canara-pdf',
        _learned: !!learned[cp.toUpperCase()], _balance: r.balance });
    });
    return { fresh: out, duplicates: dupes.length };
  }

  const api = { pdfToLines, parseCanara, prepare, counterparty, ruleCategory };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Statement = api;
})(typeof window !== 'undefined' ? window : globalThis);
