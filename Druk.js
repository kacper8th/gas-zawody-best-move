/**
 * Wydruk wyników do PDF: wszystkie kategorie w jednym pliku, każda od nowej strony.
 * Korzysta z escapeHtml_() i downloadDialogHtml_() z Regulamin.js.
 */

function printResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cfg = readConfig_(ss);

  const order = cfg.categories.map(function (c) { return c.name; });
  const rank = function (x) { const i = order.indexOf(x.meta.category); return i < 0 ? 999 : i; };
  const lists = getListSheets_(ss).sort(function (a, b) { return rank(a) - rank(b); });

  // Wyniki + szybka kontrola przed drukiem.
  let problems = 0;
  const sections = [];
  lists.forEach(function (x) {
    const res = categoryResult_(cfg, x.sheet, x.meta);
    const values = x.sheet.getRange(LIST.firstRow, 1, LIST.maxRows, x.meta.tieCol).getValues();
    problems += validateList_(values, x.meta, cfg).errors.length + res.unresolved.length;
    if (!res.rows.length) return;
    sections.push({
      category: x.meta.category,
      comps: x.meta.comps.map(function (c) {
        const live = cfg.competitions.filter(function (k) { return k.key === c.key; })[0];
        return (live || c).name;
      }),
      rows: res.rows,
      unresolved: res.unresolved,
    });
  });

  if (!sections.length) throw new Error('Brak zawodników na listach startowych - nie ma czego drukować.');

  if (problems) {
    const ans = ui.alert('Drukuj wyniki',
      'Walidacja znalazła ' + problems + ' problem(ów) (błędy wpisów lub nierozstrzygnięte remisy).\n' +
      'Szczegóły: 🏋️ Zawody → ✅ Walidacja zawodów.\n\nDrukować mimo to?', ui.ButtonSet.YES_NO);
    if (ans !== ui.Button.YES) return;
  }

  const title = cfg.title || 'Wyniki zawodów';
  const fileName = ('Wyniki - ' + title).replace(/[\\/:*?"<>|]/g, '-').slice(0, 120) + '.pdf';
  const stamp = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'dd.MM.yyyy HH:mm');

  const html = resultsHtml_(title, sections, cfg.decimals, stamp);
  const pdf = Utilities.newBlob(html, 'text/html', 'wyniki.html').getAs('application/pdf').setName(fileName);
  const dialog = HtmlService.createHtmlOutput(downloadDialogHtml_(Utilities.base64Encode(pdf.getBytes()), fileName))
    .setWidth(420).setHeight(170);
  ui.showModalDialog(dialog, '🖨️ Wyniki (PDF)');
}

/** Lista zawodników do PDF: każda kategoria od nowej strony (Nr, Imię, Nazwisko, Waga). */
function printAthletes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = readConfig_(ss);
  const order = cfg.categories.map(function (c) { return c.name; });
  const rank = function (x) { const i = order.indexOf(x.meta.category); return i < 0 ? 999 : i; };
  const lists = getListSheets_(ss).sort(function (a, b) { return rank(a) - rank(b); });

  const sections = [];
  lists.forEach(function (x) {
    const values = x.sheet.getRange(LIST.firstRow, 1, LIST.maxRows, LIST.fixedCols).getValues();
    const rows = values.filter(function (r) { return String(r[1]).trim() || String(r[2]).trim(); });
    if (!rows.length) return;
    const comps = x.meta.comps.map(function (c) {
      const live = cfg.competitions.filter(function (k) { return k.key === c.key; })[0];
      return (live || c).name;
    });
    sections.push({ category: x.meta.category, comps: comps, rows: rows });
  });
  if (!sections.length) throw new Error('Brak zawodników na listach startowych - nie ma czego drukować.');

  const title = cfg.title || 'Lista zawodników';
  const fileName = ('Lista zawodników - ' + title).replace(/[\\/:*?"<>|]/g, '-').slice(0, 120) + '.pdf';
  const stamp = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'dd.MM.yyyy HH:mm');
  const pdf = Utilities.newBlob(athletesHtml_(title, sections, stamp), 'text/html', 'lista.html')
    .getAs('application/pdf').setName(fileName);
  const dialog = HtmlService.createHtmlOutput(downloadDialogHtml_(Utilities.base64Encode(pdf.getBytes()), fileName))
    .setWidth(420).setHeight(170);
  SpreadsheetApp.getUi().showModalDialog(dialog, '🖨️ Lista zawodników (PDF)');
}

function athletesHtml_(title, sections, stamp) {
  const accent = '#1f4e79';
  const css = [
    '@page { size: A4; margin: 14mm 14mm; }',
    'body { font-family: Arial, Helvetica, sans-serif; font-size: 10.5pt; color: #222; margin: 0; }',
    '.page { page-break-after: always; }',
    '.page.last { page-break-after: auto; }',
    '.event { font-size: 11pt; color: #555; text-align: center; margin: 0 0 2pt; }',
    '.cat { font-size: 20pt; font-weight: bold; color: ' + accent + '; text-align: center; margin: 0 0 4pt; }',
    '.comps { font-size: 10pt; color: #555; text-align: center; margin: 0 0 12pt; padding-bottom: 6pt; border-bottom: 2pt solid ' + accent + '; }',
    'table { width: 100%; border-collapse: collapse; }',
    'th, td { border: 0.75pt solid #c0c7d4; padding: 5pt 6pt; }',
    'th { background: #d9e2f3; text-align: center; }',
    'td.c { text-align: center; width: 14%; }',
    'tr:nth-child(even) td { background: #f5f7fb; }',
    '.foot { margin-top: 10pt; font-size: 8pt; color: #888; text-align: right; }',
  ].join('\n');

  const pages = sections.map(function (s, si) {
    const body = s.rows.map(function (r) {
      return '<tr><td class="c">' + escapeHtml_(r[0]) + '</td><td>' + escapeHtml_(String(r[1]).trim()) + '</td>' +
        '<td>' + escapeHtml_(String(r[2]).trim()) + '</td><td class="c">' + fmtPl_(toNum_(r[LIST.weightCol - 1]), 1, false) + '</td></tr>';
    }).join('');
    return '<div class="page' + (si === sections.length - 1 ? ' last' : '') + '">' +
      '<div class="event">' + escapeHtml_(title) + '</div>' +
      '<div class="cat">' + escapeHtml_(s.category) + '</div>' +
      '<div class="comps">' + escapeHtml_(s.comps.join(' · ')) + ' · zawodników: ' + s.rows.length + '</div>' +
      '<table><thead><tr><th>Nr</th><th>Imię</th><th>Nazwisko</th><th>Waga [kg]</th></tr></thead><tbody>' + body + '</tbody></table>' +
      '<div class="foot">Wydruk: ' + escapeHtml_(stamp) + '</div></div>';
  }).join('\n');

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' + css + '</style></head><body>' + pages + '</body></html>';
}

/** Liczba po polsku: przecinek dziesiętny. fixed = stała liczba miejsc, inaczej max miejsc bez zer na końcu. */
function fmtPl_(v, decimals, fixed) {
  if (v === null || v === undefined || v === '') return '';
  const s = fixed ? Number(v).toFixed(decimals) : String(Number(Number(v).toFixed(decimals)));
  return s.replace('.', ',');
}

/** HTML wyników do konwersji na PDF (A4, kategoria = osobna strona). */
function resultsHtml_(title, sections, decimals, stamp) {
  const accent = '#1f4e79';
  const css = [
    '@page { size: A4; margin: 14mm 12mm; }',
    'body { font-family: Arial, Helvetica, sans-serif; font-size: 9.5pt; color: #222; margin: 0; }',
    '.page { page-break-after: always; }',
    '.page.last { page-break-after: auto; }',
    '.event { font-size: 11pt; color: #555; text-align: center; margin: 0 0 2pt; }',
    '.cat { font-size: 20pt; font-weight: bold; color: ' + accent + '; text-align: center; margin: 0 0 12pt; padding-bottom: 6pt; border-bottom: 2pt solid ' + accent + '; }',
    'table { width: 100%; border-collapse: collapse; }',
    'th, td { border: 0.75pt solid #c0c7d4; padding: 4pt 3pt; text-align: center; }',
    'th { background: #d9e2f3; font-size: 8.5pt; }',
    'th.grp { background: ' + accent + '; color: #fff; font-size: 9.5pt; }',
    'td.name { text-align: left; }',
    'td.place { font-weight: bold; font-size: 11pt; }',
    'td.pts { font-weight: bold; }',
    '.sep { border-left: 2pt solid ' + accent + '; }',
    'tr.p1 td { background: #fff2cc; } tr.p2 td { background: #eeeeee; } tr.p3 td { background: #fce5cd; }',
    '.warn { color: #c00000; font-weight: bold; margin-top: 8pt; }',
    '.foot { margin-top: 10pt; font-size: 8pt; color: #888; text-align: right; }',
  ].join('\n');

  const pages = sections.map(function (s, si) {
    let head1 = '<tr><th rowspan="2">Miejsce</th><th rowspan="2">Zawodnik</th><th rowspan="2">Waga [kg]</th>';
    let head2 = '<tr>';
    s.comps.forEach(function (name) {
      head1 += '<th class="grp sep" colspan="3">' + escapeHtml_(name) + '</th>';
      head2 += '<th class="sep">Najlepsza [kg]</th><th>Wynik</th><th>Miejsce</th>';
    });
    head1 += '<th class="grp sep" colspan="3">Podsumowanie</th></tr>';
    head2 += '<th class="sep">Suma miejsc</th><th>Dogrywka</th><th>Punkty</th></tr>';

    const body = s.rows.map(function (a) {
      const cls = a.place >= 1 && a.place <= 3 ? ' class="p' + a.place + '"' : '';
      let tr = '<tr' + cls + '><td class="place">' + a.place + '</td><td class="name">' + escapeHtml_(a.name) + '</td>' +
        '<td>' + fmtPl_(a.bw, 1, false) + '</td>';
      a.comps.forEach(function (c) {
        tr += '<td class="sep">' + fmtPl_(c.best, 1, false) + '</td><td>' + fmtPl_(c.score, decimals, true) + '</td><td>' + c.place + '</td>';
      });
      tr += '<td class="sep">' + a.sum + '</td><td>' + (a.tiebreak ? fmtPl_(a.tiebreak, 2, false) : '') + '</td>' +
        '<td class="pts">' + fmtPl_(a.points, 2, false) + '</td></tr>';
      return tr;
    }).join('');

    const warn = s.unresolved.map(function (u) {
      return '<div class="warn">⚠ Nierozstrzygnięty remis na miejscu ' + u.place + ': ' + escapeHtml_(u.names.join(', ')) + '</div>';
    }).join('');

    return '<div class="page' + (si === sections.length - 1 ? ' last' : '') + '">' +
      '<div class="event">' + escapeHtml_(title) + '</div>' +
      '<div class="cat">' + escapeHtml_(s.category) + '</div>' +
      '<table><thead>' + head1 + head2 + '</thead><tbody>' + body + '</tbody></table>' + warn +
      '<div class="foot">Wydruk: ' + escapeHtml_(stamp) + '</div></div>';
  }).join('\n');

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' + css + '</style></head><body>' + pages + '</body></html>';
}
