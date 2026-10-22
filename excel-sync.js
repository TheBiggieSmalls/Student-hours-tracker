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
      if (seen[k]) rep.skipped.push(nlDate(k) + ' staat al in je Excel'); else fresh.push(e);
    });

    var L = ticketLines(ws, b), srcRow = lastFilled || b.h + 1;
    var row = (lastFilled || b.h) + 1, lastWritten = lastFilled;
    var add = {}, singlePrices = [];     // teller per ticketregel
    function bump(rowNr, n) { add[rowNr] = (add[rowNr] || 0) + n; }
    var lastRowAllowed = b.s - 1;

    fresh.forEach(function (e) {
      var k = dk(e.start), loc = e.loc || '', single = e.trip === 'single', home = !!(locs[loc] && locs[loc].home), nm = loc.trim().toLowerCase();
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

  var api = { run: run, lastSheetDate: lastSheetDate, dk: dk, eur: eur, nlDate: nlDate };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.WerkurenExcel = api;
})(typeof window !== 'undefined' ? window : this);
