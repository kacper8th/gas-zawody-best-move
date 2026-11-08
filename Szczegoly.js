/**
 * Zakładka „Wyniki szczegółowe”: dla każdej kategorii osobny ranking każdej konkurencji,
 * posortowany wg miejsca w tej konkurencji (próby z zaliczeniem, najlepsza, wynik).
 * Zakładka wyliczana - przebudowywana razem z „Wyniki”.
 */

const SHEET_DETAILS = 'Wyniki szczegółowe';

function getDetailsSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_DETAILS);
  if (sh) return sh;
  const results = ss.getSheetByName(SHEET_RESULTS);
  sh = ss.insertSheet(SHEET_DETAILS, results ? results.getIndex() : ss.getSheets().length);
  sh.protect().setDescription('zbm:szczegoly').setWarningOnly(true);
  sh.setColumnWidth(1, 70);
  sh.setColumnWidth(2, 220);
  return sh;
}

/** Tekst próby: „18,0 ✓” (zaliczona) / „20,0 ✗” (spalona) / '' (brak). */
function attemptLabel_(a) {
  if (!a || a.kg === null) return '';
  if (a.kg === 0 && !a.ok) return '✗'; // spalona, ciężar nieznany
  return String(Number(a.kg).toFixed(1)).replace('.', ',') + (a.ok ? ' ✓' : ' ✗');
}

/**
 * Czysta funkcja: wiersze i formatowanie zakładki.
 * collected: [{ meta, res }] (res z computeRanking_, wiersze z polem raw = próby per konkurencja).
 * Zwraca { rows, fmt } - fmt: { row, kind, width?, count?, scoreCol?, bestCol? }.
 */
function detailsLayout_(title, stamp, collected, competitions) {
  const rows = [], fmt = [];
  rows.push([title ? 'Wyniki szczegółowe - ' + title : 'Wyniki szczegółowe']);
  fmt.push({ row: rows.length, kind: 'title' });
  rows.push(['Przeliczono: ' + stamp]);
  fmt.push({ row: rows.length, kind: 'info' });

  if (!collected.length) {
    rows.push(['']);
    rows.push(['Brak list startowych - ⚙️ Setup → Generuj listy startowe.']);
    fmt.push({ row: rows.length, kind: 'info' });
  }

  collected.forEach(function (x) {
    rows.push(['']);
    rows.push([x.meta.category]);
    fmt.push({ row: rows.length, kind: 'section' });

    x.meta.comps.forEach(function (c, ci) {
      const live = competitions.filter(function (k) { return k.key === c.key; })[0];
      const header = ['Miejsce', 'Zawodnik', 'Waga [kg]'];
      for (let i = 1; i <= c.attempts; i++) header.push('Próba ' + i);
      header.push('Najlepsza [kg]', 'Wynik');
      const width = header.length;

      rows.push([(live || c).name]);
      fmt.push({ row: rows.length, kind: 'comp', width: width });
      rows.push(header);
      fmt.push({ row: rows.length, kind: 'header', width: width });

      const sorted = x.res.rows.slice().sort(function (a, b) {
        return a.comps[ci].place - b.comps[ci].place || (Number(a.nr) || 0) - (Number(b.nr) || 0);
      });
      if (!sorted.length) {
        rows.push(['', '(brak zawodników)']);
        fmt.push({ row: rows.length, kind: 'info' });
      } else {
        const first = rows.length + 1;
        sorted.forEach(function (a) {
          const cc = a.comps[ci];
          const raw = (a.raw && a.raw[ci]) || [];
          const r = [cc.place, a.name, a.bw === null ? '' : a.bw];
          for (let i = 0; i < c.attempts; i++) r.push(attemptLabel_(raw[i]));
          r.push(cc.best === null ? '' : cc.best, cc.score === null ? '' : cc.score);
          rows.push(r);
        });
        fmt.push({ row: first, kind: 'body', count: sorted.length, width: width, bestCol: width - 1, scoreCol: width,
          // Kolor podium tylko dla zawodników z wynikiem (bez wyniku = 0 = bez koloru).
          places: sorted.map(function (a) { return a.comps[ci].score === null ? 0 : a.comps[ci].place; }) });
      }
      rows.push(['']);
    });
  });
  return { rows: rows, fmt: fmt };
}

/** Przebudowa zakładki (wywoływana z rebuildResultsUnlocked_). */
function buildDetailsSheet_(ss, cfg, collected, scoreFmt) {
  const sh = getDetailsSheet_(ss);
  const stamp = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'yyyy-MM-dd HH:mm:ss');
  const layout = detailsLayout_(cfg.title, stamp, collected, cfg.competitions);
  const width = layout.rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
  const grid = layout.rows.map(function (r) { return r.concat(new Array(width - r.length).fill('')); });

  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  if (sh.getMaxRows() < grid.length) sh.insertRowsAfter(sh.getMaxRows(), grid.length - sh.getMaxRows());
  sh.getRange(1, 1, grid.length, width).setValues(grid).setVerticalAlignment('middle');

  const podium = ['#fff2cc', '#eeeeee', '#fce5cd'];
  layout.fmt.forEach(function (f) {
    if (f.kind === 'title') sh.getRange(f.row, 1).setFontSize(16).setFontWeight('bold');
    if (f.kind === 'info') sh.getRange(f.row, 1, 1, 2).setFontColor('#666666').setFontStyle('italic');
    if (f.kind === 'section') sh.getRange(f.row, 1).setFontSize(13).setFontWeight('bold');
    if (f.kind === 'comp') {
      sh.getRange(f.row, 1, 1, f.width).merge()
        .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff').setHorizontalAlignment('center');
    }
    if (f.kind === 'header') {
      sh.getRange(f.row, 1, 1, f.width).setFontWeight('bold').setBackground('#d9e2f3').setWrap(true).setHorizontalAlignment('center');
    }
    if (f.kind === 'body') {
      const body = sh.getRange(f.row, 1, f.count, f.width);
      body.setHorizontalAlignment('center')
        .setBorder(true, true, true, true, true, true, '#c0c7d4', SpreadsheetApp.BorderStyle.SOLID);
      sh.getRange(f.row, 2, f.count, 1).setHorizontalAlignment('left');
      sh.getRange(f.row, 1, f.count, 1).setFontWeight('bold');
      sh.getRange(f.row, 3, f.count, 1).setNumberFormat('0.0');
      sh.getRange(f.row, f.bestCol, f.count, 1).setNumberFormat('0.0').setFontWeight('bold');
      sh.getRange(f.row, f.scoreCol, f.count, 1).setNumberFormat(scoreFmt);
      f.places.forEach(function (p, i) {
        if (p >= 1 && p <= 3) sh.getRange(f.row + i, 1, 1, f.width).setBackground(podium[p - 1]);
      });
    }
  });
}
