/* Werkuren -> Excel: vult nieuwe werkdagen in je bestaande uren-Excel in. */
(function (root) {
  var KAART = 'leuven';            // plaats die je met de 10-ritkaart doet (regel "10 rit kaart")
  var TIME_FMT = '[h]:mm', DATE_FMT = 'dd/mm/yyyy';

  function p2(n) { return n < 10 ? '0' + n : '' + n; }
  function key(y, m, d) { return y + '-' + p2(m) + '-' + p2(d); }
  function dk(ms) { var d = new Date(ms); return key(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
  function frac(ms) { var d = new Date(ms); return (d.getHours() * 60 + d.getMinutes()) / 1440; }
  function eur(n) { return '€ ' + (Math.round(n * 100) / 100).toFixed(2).replace('.', ','); }
  function nlDate(k) { var a = k.split('-'); return a[2] + '/' + a[1] + '/' + a[0]; }
  function cv(cell) {
    var v = cell.value;
    if (v && typeof v === 'object' && !(v instanceof Date)) {
      if ('result' in v) return v.result;
      if (v.richText) return v.richText.map(function (t) { return t.text; }).join('');
      if ('text' in v) return v.text;
    }
    return v;
  }
  function lc(cell) { var v = cv(cell); return typeof v === 'string' ? v.trim().toLowerCase() : ''; }
  function hasFormula(cell) { var v = cell.value; return !!(v && typeof v === 'object' && v.formula); }

  // 'y-m-d' kandidaten voor een datumcel (beste eerst; dag/maand omgewisseld als tweede)
  function dateCands(v) {
    var out = [], y, m, d, mt;
    if (v instanceof Date) { y = v.getUTCFullYear(); m = v.getUTCMonth() + 1; d = v.getUTCDate(); }
    else if (typeof v === 'string' && (mt = v.trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/))) { d = +mt[1]; m = +mt[2]; y = +mt[3]; if (y < 100) y += 2000; }
    else return out;
    out.push(key(y, m, d)); if (d <= 12 && d !== m) out.push(key(y, d, m));
    return out;
  }

  function findBlocks(ws) {
    var hs = [];
    ws.eachRow({ includeEmpty: false }, function (row, n) { var v = cv(row.getCell(2)); if (typeof v === 'string' && v.trim().toLowerCase() === 'datum') hs.push(n); });
    return hs.map(function (h, i) {
      var end = (hs[i + 1] || ws.rowCount + 1) - 1, s = null, tot = null;
      for (var r = h + 1; r <= end; r++) {
        var f = cv(ws.getCell(r, 6)); if (s === null && typeof f === 'string' && /^betaling/i.test(f.trim())) s = r;
        var c = cv(ws.getCell(r, 3)); if (typeof c === 'string' && /totaal ritten/i.test(c)) tot = r;
      }
      return { h: h, s: s, end: end, tot: tot };
    });
  }
  function ticketLines(ws, b) {
    var L = {};
    for (var r = b.s + 1; r <= (b.tot || b.end); r++) {
      var a = lc(ws.getCell(r, 1)), bb = lc(ws.getCell(r, 2));
      if (a.indexOf('10 rit') === 0) { if (bb.indexOf('heen') === 0) L.kRt = r; else if (bb.indexOf('rit enkel') === 0) L.kSingle = r; }
      else if (a === 'ticket') { if (bb === 'ander rit') L.other = r; else if (bb) L['t:' + bb] = r; }
    }
    return L;
  }
  function priceAt(prices, loc, k) {
    var list = (prices && prices[loc]) || [], best = null;
    list.forEach(function (p) { if (p.from <= k && (!best || p.from > best.from)) best = p; });
    return best || (list.length ? list.slice().sort(function (a, b) { return a.from < b.from ? -1 : 1; })[0] : null);
  }
  function copyStyle(ws, srcRow, dstRow, col) {
    var s = ws.getCell(srcRow, col), d = ws.getCell(dstRow, col);
    if (s.style && Object.keys(s.style).length) d.style = JSON.parse(JSON.stringify(s.style));
  }
  function extendRange(cell, lastRow) {
    var v = cell.value; if (!v || !v.formula) return false;
    var changed = false, f = v.formula.replace(/([A-Z]+)(\d+):([A-Z]+)(\d+)/, function (all, c1, r1, c2, r2) {
      if (+r2 < lastRow) { changed = true; return c1 + r1 + ':' + c2 + lastRow; } return all;
    });
    if (changed) cell.value = { formula: f }; return changed;
  }

  // Zoek de laatste datum in je Excel (voor de standaard "van"-datum)
  function lastSheetDate(wb) {
    var ws = wb.worksheets[0], best = null;
    findBlocks(ws).forEach(function (b) {
      var end = (b.s || b.end) - 1;
      for (var r = b.h + 1; r <= end; r++) { var c = dateCands(cv(ws.getCell(r, 2)))[0]; if (c && (!best || c > best)) best = c; }
    });
    return best;
  }

  function run(wb, data, opts) {
    var ws = wb.worksheets[0], rep = { added: [], skipped: [], tickets: [], warnings: [], nowjobs: [], block: null };
    var blocks = findBlocks(ws);
    if (!blocks.length) { rep.warnings.push('Ik vind geen tabel met een kolom "Datum" in je Excel.'); return rep; }
    var b = blocks[blocks.length - 1];
    if (!b.s) { rep.warnings.push('Ik vind de regel "Betaling +2euro/uur" niet onder het laatste blok.'); return rep; }
    rep.block = b;

    // reeds ingevulde datums (alle blokken)
    var seen = {}, lastFilled = 0;
    blocks.forEach(function (bl) {
      for (var r = bl.h + 1; r <= (bl.s || bl.end) - 1; r++) {
        var v = cv(ws.getCell(r, 2)); if (v === null || v === undefined || v === '') continue;
        dateCands(v).forEach(function (c) { seen[c] = 1; });
        if (bl === b) lastFilled = r;
      }
    });

    var locs = {}; (data.locs || []).forEach(function (l) { locs[l.n] = l; });
    var from = opts.from || '0000-00-00', to = opts.to || '9999-99-99';
    var todo = (data.entries || []).filter(function (e) { var k = dk(e.start); return e.end != null && k >= from && k <= to; })
      .sort(function (a, c) { return a.start - c.start; });
    var fresh = [];
    todo.forEach(function (e) {
      var k = dk(e.start);
      if (seen[k] && !opts.allowDup) rep.skipped.push(nlDate(k) + ' staat al in je Excel'); else fresh.push(e);
    });

    var L = ticketLines(ws, b), srcRow = lastFilled || b.h + 1;
    var row = (lastFilled || b.h) + 1, lastWritten = lastFilled;
    var add = {}, singlePrices = [];     // teller per ticketregel
    function bump(rowNr, n) { add[rowNr] = (add[rowNr] || 0) + n; }
    var lastRowAllowed = b.s - 1;

    fresh.forEach(function (e) {
      var k = dk(e.start), loc = e.loc || '', single = e.trip === 'single', home = !!(locs[loc] && locs[loc].home) || e.trip === 'none', nm = loc.trim().toLowerCase();
      if (row > lastRowAllowed) { rep.skipped.push(nlDate(k) + ' past niet meer in dit blok (voeg zelf rijen toe)'); return; }
      // dagrij
      ['A', 'B', 'C', 'D', 'F', 'I', 'J'].forEach(function (c, i) { copyStyle(ws, srcRow, row, 'ABCDEFGHIJ'.indexOf(c) + 1); });
      ws.getCell(row, 1).value = loc + (single && !home ? ' enkel' : '');
      var bc = ws.getCell(row, 2); bc.value = new Date(Date.UTC(+k.slice(0, 4), +k.slice(5, 7) - 1, +k.slice(8, 10))); bc.numFmt = DATE_FMT;
      var hasNj = !!(e.njS && e.njE);
      function hmFrac(t) { var a = t.split(':'); return (+a[0] * 60 + +a[1]) / 1440; }
      [[3, frac(e.start)], [4, frac(e.end)], [9, hasNj ? hmFrac(e.njS) : frac(e.start)], [10, hasNj ? hmFrac(e.njE) : frac(e.end)]].forEach(function (p) { var c = ws.getCell(row, p[0]); c.value = p[1]; c.numFmt = TIME_FMT; });
      if (hasNj) {
        var njm = hmFrac(e.njS) * 1440, nje = hmFrac(e.njE) * 1440, nj = nje - njm; if (nj < 0) nj += 1440;
        var diff = Math.round((e.end - e.start) / 60000) - Math.round(nj);
        rep.nowjobs.push('Rij ' + row + ' · ' + nlDate(k) + ': NowJobs ' + e.njS + '–' + e.njE + ', echt ' + p2(new Date(e.start).getHours()) + ':' + p2(new Date(e.start).getMinutes()) + '–' + p2(new Date(e.end).getHours()) + ':' + p2(new Date(e.end).getMinutes()) + ' → ' + (diff > 0 ? diff + ' min extra te betalen' : diff < 0 ? 'NowJobs telt ' + (-diff) + ' min meer' : 'geen verschil'));
      }
      var fm = { 5: 'SUM(D' + row + '-C' + row + ')', 7: 'SUM((E' + row + '*24)*2)', 11: 'SUM(J' + row + '-I' + row + ')', 12: 'SUM(E' + row + '-K' + row + ')', 13: 'SUM((24*L' + row + ')*16)' };
      Object.keys(fm).forEach(function (col) { var c = ws.getCell(row, +col); if (!hasFormula(c)) { c.value = { formula: fm[col] }; copyStyle(ws, srcRow, row, +col); } });
      var mf = ws.getCell(row, 13).value; if (mf && mf.formula && /\*-16/.test(mf.formula)) rep.warnings.push('Rij ' + row + ': kolom "Te betalen uren" gebruikt −16 in plaats van 16 (zo gelaten zoals in je sjabloon).');
      var tr = home ? '' : (single ? 'enkel' : 'heen en terug');
      var pr = home ? null : priceAt(data.prices, loc, k);

      // trein
      if (!home) {
        if (nm === KAART) {
          if (single) { ws.getCell(row, 6).value = 'X'; rep.tickets.push('Rij ' + row + ': "X" in Treinrit enkel (10-ritkaart, enkele rit)'); }
          else if (L.kRt) bump(L.kRt, 1); else rep.warnings.push('Geen regel "10 rit kaart – Heen & terug" gevonden voor ' + nlDate(k) + '.');
        } else if (single) {
          if (L.other) { bump(L.other, 1); singlePrices.push(pr ? pr.single : null); } else rep.warnings.push('Geen regel "Ander rit" gevonden voor ' + nlDate(k) + ' (' + loc + ' enkel).');
        } else if (L['t:' + nm]) bump(L['t:' + nm], 1);
        else rep.warnings.push('Geen ticketregel voor "' + loc + '" in je Excel (' + nlDate(k) + '). Voeg die regel zelf toe.');
        if (!pr) rep.warnings.push('Geen ticketprijs in de app voor ' + loc + ' op ' + nlDate(k) + '.');
      }
      rep.added.push({ row: row, date: k, start: e.start, end: e.end, loc: loc, trip: tr });
      lastWritten = row; row++;
    });

    // ticketregels bijwerken
    Object.keys(add).forEach(function (r) {
      r = +r; var dC = ws.getCell(r, 4), cC = ws.getCell(r, 3), label = (cv(ws.getCell(r, 2)) || '').toString().trim();
      if (hasFormula(dC) || (dC.value !== null && dC.value !== undefined && typeof dC.value !== 'number')) { rep.warnings.push('Aantal bij "' + label + '" (D' + r + ') is geen vast getal; niet aangepast.'); return; }
      var old = dC.value || 0; dC.value = old + add[r];
      rep.tickets.push('D' + r + ' (' + label + '): ' + old + ' → ' + (old + add[r]));
      if (r === L.other) {
        var sp = singlePrices.filter(function (x) { return x !== null; });
        if (sp.length) {
          var oldP = typeof cC.value === 'number' ? cC.value : 0, tot = oldP * old + sp.reduce(function (a, x) { return a + x; }, 0), avg = Math.round(tot / (old + sp.length) * 10000) / 10000;
          if (Math.abs(avg - oldP) > 0.0001) { cC.value = avg; rep.tickets.push('C' + r + ' (Ander rit): gemiddelde prijs ' + eur(oldP) + ' → ' + eur(avg) + ' zodat het totaal klopt'); }
        }
      }
    });
    // prijs-controle voor heen&terug-regels
    var checked = {};
    rep.added.forEach(function (a) {
      if (a.trip !== 'heen en terug') return;
      var nm = a.loc.trim().toLowerCase(), r = nm === KAART ? L.kRt : L['t:' + nm]; if (!r || checked[r]) return; checked[r] = 1;
      var pr = priceAt(data.prices, a.loc, a.date), cC = ws.getCell(r, 3);
      if (pr && typeof cC.value === 'number' && Math.abs(cC.value - pr.rt) > 0.005) {
        var oldD = (add[r] ? ws.getCell(r, 4).value - add[r] : ws.getCell(r, 4).value) || 0;
        if (oldD === 0) { rep.tickets.push('C' + r + ' (' + a.loc + '): prijs ' + eur(cC.value) + ' → ' + eur(pr.rt) + ' (nieuwe prijs uit de app)'); cC.value = pr.rt; }
        else rep.warnings.push('Prijs ' + a.loc + ': Excel heeft ' + eur(cC.value) + ', de app ' + eur(pr.rt) + '. Niet aangepast omdat er al ritten in dit blok staan.');
      }
    });
    // sommen laten meelopen tot de laatste rij
    if (lastWritten) {
      var ext = [ws.getCell(b.s, 7), ws.getCell(b.s, 13)]; if (L.kSingle) ext.push(ws.getCell(L.kSingle, 4));
      ext.forEach(function (c) { if (extendRange(c, lastWritten)) rep.tickets.push(c.address + ': som uitgebreid tot rij ' + lastWritten); });
    }
    wb.calcProperties = wb.calcProperties || {}; wb.calcProperties.fullCalcOnLoad = true;
    return rep;
  }


  // ---- Excel-tab in de app: lees, bereken en bewerk het laatste blok
  function tmin(v) {
    if (v instanceof Date) return v.getUTCHours() * 60 + v.getUTCMinutes();
    if (typeof v === 'number') return Math.round((v % 1) * 1440) % 1440;
    return null;
  }
  function hhmm(m) { return p2(Math.floor(m / 60)) + ':' + p2(m % 60); }
  function ftext(cell) { var v = cell.value; return v && typeof v === 'object' ? (v.formula || '') : ''; }
  function rateOf(cell, def) { var m = ftext(cell).match(/\*\s*(-?\d+(?:\.\d+)?)\)+\s*$/); return m ? parseFloat(m[1]) : def; }
  function r2(x) { return Math.round(x * 100) / 100; }

  function titleRowOf(ws, b) { var t = b.h > 1 ? cv(ws.getCell(b.h - 1, 1)) : null; return (typeof t === 'string' && t.trim()) ? b.h - 1 : b.h; }
  function blockCount(wb) { return findBlocks(wb.worksheets[0]).length; }
  function snapshot(wb, idx) {
    var ws = wb.worksheets[0], out = { ok: false, rows: [], lines: [] }, bl = findBlocks(ws);
    var b = bl[idx === undefined || idx === null ? bl.length - 1 : idx];
    if (!b || !b.s) { out.err = 'Ik vind geen blok met "Datum" en "Betaling" in dit Excel-bestand.'; return out; }
    var rows = [], lastFilled = 0;
    for (var r = b.h + 1; r <= b.s - 1; r++) {
      var bv = cv(ws.getCell(r, 2)); if (bv === null || bv === undefined || bv === '') continue;
      lastFilled = r; rows.push({ r: r, c: dateCands(bv) });
    }
    var firsts = rows.filter(function (x) { return x.c.length; }).map(function (x) { var a = x.c[0].split('-'); return Date.UTC(+a[0], +a[1] - 1, +a[2]); }).sort(function (a, c) { return a - c; });
    var ref = firsts.length ? firsts[Math.floor(firsts.length / 2)] : 0, gSum = 0, mSum = 0, mins = 0, singles = 0, last = '';
    rows.forEach(function (x) {
      var row = { r: x.r, k: null, fixed: false };
      if (x.c.length) {
        var best = x.c[0], bd = 1e18;
        x.c.forEach(function (k) { var a = k.split('-'), dd = Math.abs(Date.UTC(+a[0], +a[1] - 1, +a[2]) - ref); if (dd < bd) { bd = dd; best = k; } });
        row.k = best; row.fixed = best !== x.c[0]; if (!last || best > last) last = best;
      }
      var cs = tmin(cv(ws.getCell(x.r, 3))), ce = tmin(cv(ws.getCell(x.r, 4))), is = tmin(cv(ws.getCell(x.r, 9))), je = tmin(cv(ws.getCell(x.r, 10)));
      row.cs = cs; row.ce = ce;
      var E = (cs !== null && ce !== null) ? ce - cs + (ce < cs ? 1440 : 0) : 0, K = (is !== null && je !== null) ? je - is + (je < is ? 1440 : 0) : 0;
      var lf = ftext(ws.getCell(x.r, 12)), L = /K\d+\s*-\s*E\d+/.test(lf) ? K - E : E - K;
      var g = E / 60 * rateOf(ws.getCell(x.r, 7), 2), m = L / 60 * rateOf(ws.getCell(x.r, 13), 16);
      row.mins = E; row.g = g; row.m = m;
      if (cs !== null && ce !== null && is !== null && je !== null && (is !== cs || je !== ce)) { row.njS = hhmm(is); row.njE = hhmm(je); row.diff = E - K; }
      var f = cv(ws.getCell(x.r, 6)); row.x = f !== null && f !== undefined && String(f).trim() !== ''; if (row.x) singles++;
      var a = (cv(ws.getCell(x.r, 1)) || '').toString().trim(); row.single = /enkel/i.test(a) || row.x; row.loc = a.replace(/\s*enkel\s*/i, ' ').trim();
      gSum += g; if (m > 0) mSum += m; mins += E;
      out.rows.push(row);
    });
    var L = ticketLines(ws, b), train = 0;
    function line(key, kind) {
      var rn = L[key]; if (!rn) return;
      var label = ((cv(ws.getCell(rn, 1)) || '') + ' ' + (cv(ws.getCell(rn, 2)) || '')).toString().replace(/\s+/g, ' ').trim();
      var price = cv(ws.getCell(rn, 3)), cnt = cv(ws.getCell(rn, 4)); price = typeof price === 'number' ? price : null; cnt = typeof cnt === 'number' ? cnt : 0;
      var total = kind === 'kRt' ? (price || 0) * cnt + (price || 0) * singles / 2 : (price || 0) * cnt;
      train += total; out.lines.push({ r: rn, kind: kind, label: label, price: price, count: cnt, total: r2(total), editable: true });
    }
    line('kRt', 'kRt');
    if (L.kSingle) out.lines.push({ r: L.kSingle, kind: 'kSingle', label: '10 rit kaart Rit enkel', price: null, count: singles, total: 0, editable: false });
    Object.keys(L).filter(function (k) { return k.indexOf('t:') === 0 || k === 'other'; }).sort(function (a, c) { return L[a] - L[c]; }).forEach(function (k) { line(k, 't'); });
    var carry = 0, sun = 0, carryRow = null, sunRow = null;
    for (var q = b.s; q <= b.end; q++) {
      var lab = lc(ws.getCell(q, 6)), gv = cv(ws.getCell(q, 7));
      if (/vorige keer/.test(lab)) { carryRow = q; if (typeof gv === 'number') carry = gv; }
      if (/150%/.test(lab)) { sunRow = q; if (typeof gv === 'number') sun = gv; }
    }
    out.ok = true; out.titleRow = titleRowOf(ws, b); out.from = rows.length ? (out.rows.map(function (r) { return r.k; }).filter(Boolean).sort()[0] || '') : ''; out.b = b; out.mins = mins; out.g = r2(gSum); out.m = r2(mSum); out.train = r2(train);
    out.carry = carry; out.sun = sun; out.carryRow = carryRow; out.sunRow = sunRow; out.last = last; out.next = (lastFilled || b.h) + 1; out.max = b.s - 1;
    out.rates = { g: rateOf(ws.getCell(b.h + 1, 7), 2), m: 16 };
    out.total = r2(carry + gSum + mSum + train + sun);
    return out;
  }
  function pendingEntries(wb, entries) {
    var ws = wb.worksheets[0], seen = {}, last = lastSheetDate(wb) || '0000-00-00';
    findBlocks(ws).forEach(function (bl) { for (var r = bl.h + 1; r <= (bl.s || bl.end) - 1; r++) { var v = cv(ws.getCell(r, 2)); if (v === null || v === undefined || v === '') continue; dateCands(v).forEach(function (c) { seen[c] = 1; }); } });
    var days = {};
    return (entries || []).filter(function (e) { var k = dk(e.start); if (e.end == null || k <= last || seen[k] || days[k]) return false; days[k] = 1; return true; });
  }
  function ensureFormulas(ws, row, srcRow) {
    var fm = { 5: 'SUM(D' + row + '-C' + row + ')', 7: 'SUM((E' + row + '*24)*2)', 11: 'SUM(J' + row + '-I' + row + ')', 12: 'SUM(E' + row + '-K' + row + ')', 13: 'SUM((24*L' + row + ')*16)' };
    Object.keys(fm).forEach(function (col) { var c = ws.getCell(row, +col); if (!hasFormula(c)) { c.value = { formula: fm[col] }; copyStyle(ws, srcRow, row, +col); } });
  }
  function clearRow(wb, row) { var ws = wb.worksheets[0]; [1, 2, 3, 4, 6, 9, 10].forEach(function (c) { ws.getCell(row, c).value = null; }); }
  function hmFrac(t) { var a = t.split(':'); return (+a[0] * 60 + +a[1]) / 1440; }
  // sp: { k:'yyyy-mm-dd', s:'HH:MM', e:'HH:MM', ns, ne, loc, trip:'rt'|'single'|'none', home }
  function editRow(wb, row, sp) {
    var ws = wb.worksheets[0], single = sp.trip === 'single';
    ws.getCell(row, 1).value = (sp.loc || '') + (single && !sp.home ? ' enkel' : '');
    var bc = ws.getCell(row, 2); bc.value = new Date(Date.UTC(+sp.k.slice(0, 4), +sp.k.slice(5, 7) - 1, +sp.k.slice(8, 10))); bc.numFmt = DATE_FMT;
    var hasNj = !!(sp.ns && sp.ne);
    [[3, hmFrac(sp.s)], [4, hmFrac(sp.e)], [9, hasNj ? hmFrac(sp.ns) : hmFrac(sp.s)], [10, hasNj ? hmFrac(sp.ne) : hmFrac(sp.e)]].forEach(function (p) { var c = ws.getCell(row, p[0]); c.value = p[1]; c.numFmt = TIME_FMT; });
    ws.getCell(row, 6).value = (single && String(sp.loc || '').trim().toLowerCase() === KAART) ? 'X' : null;
    ensureFormulas(ws, row, row);
  }
  function setLine(wb, row, c, d) { var ws = wb.worksheets[0]; if (typeof c === 'number') ws.getCell(row, 3).value = c; if (typeof d === 'number' && !hasFormula(ws.getCell(row, 4))) ws.getCell(row, 4).value = d; }
  function setInput(wb, row, v) { wb.worksheets[0].getCell(row, 7).value = v; }


  // ---- betaling voltooid: nieuw blok onderaan, zelfde opmaak en formules, alles leeg
  function shiftFormula(f, off, minRow) {
    return f.replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g, function (m, c, r) { return +r >= minRow ? c + (+r + off) : m; });
  }
  function rowHasContent(ws, r) { for (var c = 1; c <= 15; c++) { var v = cv(ws.getCell(r, c)); if (v !== null && v !== undefined && v !== '') return true; } return false; }
  function closeBlock(wb, opts) {
    var ws = wb.worksheets[0], bl = findBlocks(ws), snap = snapshot(wb);
    if (!snap.ok) throw new Error(snap.err);
    var b = bl[bl.length - 1], titleRow = snap.titleRow, endRow = Math.min(b.end, ws.rowCount);
    while (endRow > b.h && !rowHasContent(ws, endRow)) endRow--;
    var gap = 3;
    if (bl.length > 1) { var pe = titleRow - 1; while (pe > bl[bl.length - 2].h && !rowHasContent(ws, pe)) pe--; gap = Math.max(2, Math.min(6, titleRow - pe - 1)); }
    var newTitle = endRow + gap + 1, off = newTitle - titleRow, first = b.h + 1, lastD = b.s - 1;
    var carry = r2(snap.total - (typeof opts.received === 'number' ? opts.received : snap.total));
    var L = ticketLines(ws, b), countRows = {};
    Object.keys(L).forEach(function (k) { countRows[L[k]] = k; });
    for (var r = titleRow; r <= endRow; r++) {
      var nr = r + off, isData = r >= first && r <= lastD;
      ws.getRow(nr).height = ws.getRow(r).height;
      for (var c = 1; c <= 15; c++) {
        var src = ws.getCell(r, c), dst = ws.getCell(nr, c);
        if (src.style && Object.keys(src.style).length) dst.style = JSON.parse(JSON.stringify(src.style));
        var v = src.value;
        if (isData) { dst.value = null; continue; }
        if (v && typeof v === 'object' && !(v instanceof Date)) {
          if (v.formula) dst.value = { formula: shiftFormula(v.formula, off, titleRow) };
          else dst.value = null;
        } else dst.value = (v === undefined ? null : v);
      }
      if (isData) {
        var fm = { 5: 'SUM(D' + nr + '-C' + nr + ')', 7: 'SUM((E' + nr + '*24)*2)', 11: 'SUM(J' + nr + '-I' + nr + ')', 12: 'SUM(E' + nr + '-K' + nr + ')', 13: 'SUM((24*L' + nr + ')*16)' };
        Object.keys(fm).forEach(function (col) { ws.getCell(nr, +col).value = { formula: fm[col] }; });
      }
    }
    var f1 = first + off, f2 = lastD + off;
    ws.getCell(b.s + off, 7).value = { formula: 'SUM(G' + f1 + ':G' + f2 + ')' };
    ws.getCell(b.s + off, 13).value = { formula: 'SUMIF(M' + f1 + ':M' + f2 + ',">0")' };
    if (L.kSingle) ws.getCell(L.kSingle + off, 4).value = { formula: 'COUNTA(F' + f1 + ':F' + f2 + ')' };
    Object.keys(countRows).forEach(function (rn) { rn = +rn; if (rn === L.kSingle) return; var dc = ws.getCell(rn + off, 4); if (!hasFormula(dc)) dc.value = 0; });
    if (snap.carryRow) ws.getCell(snap.carryRow + off, 7).value = carry;
    if (snap.sunRow) ws.getCell(snap.sunRow + off, 7).value = 0;
    wb.calcProperties = wb.calcProperties || {}; wb.calcProperties.fullCalcOnLoad = true;
    return { titleRow: newTitle, carry: carry, total: snap.total };
  }

  var api = { run: run, snapshot: snapshot, blockCount: blockCount, closeBlock: closeBlock, pendingEntries: pendingEntries, editRow: editRow, clearRow: clearRow, setLine: setLine, setInput: setInput, lastSheetDate: lastSheetDate, dk: dk, eur: eur, nlDate: nlDate };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.WerkurenExcel = api;
})(typeof window !== 'undefined' ? window : this);
