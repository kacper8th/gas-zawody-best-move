/**
 * Wydruk wyników szczegółowych do PDF: każda kategoria od nowej strony,
 * na stronie ranking każdej konkurencji (próby ✓/✗, najlepsza, wynik).
 * Korzysta z escapeHtml_(), fmtPl_(), attemptLabel_(), downloadDialogHtml_().
 */

function printDetails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cfg = readConfig_(ss);

  const order = cfg.categories.map(function (c) { return c.name; });
  const rank = function (x) { const i = order.indexOf(x.meta.category); return i < 0 ? 999 : i; };
  const lists = getListSheets_(ss).sort(function (a, b) { return rank(a) - rank(b); });

  const sections = [];
  lists.forEach(function (x) {
    const res = categoryResult_(cfg, x.sheet, x.meta);
    if (!res.rows.length) return;
    sections.push({
      category: x.meta.category,
      comps: x.meta.comps.map(function (c, ci) {
        const live = cfg.competitions.filter(function (k) { return k.key === c.key; })[0];
        return { name: (live || c).name, attempts: c.attempts, index: ci };
      }),
      rows: res.rows,
    });
  });
  if (!sections.length) throw new Error('Brak zawodników na listach startowych - nie ma czego drukować.');

  const v = collectValidation_(ss, cfg);
  if (v.errors.length + v.warnings.length) {
    const ans = ui.alert('Drukuj wyniki szczegółowe',
      'Walidacja znalazła problemy:\n\n' + validationReport_(v) + '\n\nDrukować mimo to?', ui.ButtonSet.YES_NO);
    if (ans !== ui.Button.YES) return;
  }

  const title = cfg.title || 'Wyniki zawodów';
  const fileName = ('Wyniki szczegółowe - ' + title).replace(/[\\/:*?"<>|]/g, '-').slice(0, 120) + '.pdf';
  const stamp = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'dd.MM.yyyy HH:mm');
  const pdf = Utilities.newBlob(detailsHtml_(title, sections, cfg.decimals, stamp), 'text/html', 'szczegoly.html')
    .getAs('application/pdf').setName(fileName);
  const dialog = HtmlService.createHtmlOutput(downloadDialogHtml_(Utilities.base64Encode(pdf.getBytes()), fileName))
    .setWidth(420).setHeight(170);
  ui.showModalDialog(dialog, '🖨️ Wyniki szczegółowe (PDF)');
}

/** HTML wyników szczegółowych (A4, kategoria = strona, konkurencje jedna pod drugą). */
function detailsHtml_(title, sections, decimals, stamp) {
  const accent = '#1f4e79';
  const css = [
    '@page { size: A4; margin: 14mm 12mm; }',
    'body { font-family: Arial, Helvetica, sans-serif; font-size: 9.5pt; color: #222; margin: 0; }',
    '.page { page-break-after: always; }',
    '.page.last { page-break-after: auto; }',
    '.event { font-size: 11pt; color: #555; text-align: center; margin: 0 0 2pt; }',
    '.cat { font-size: 20pt; font-weight: bold; color: ' + accent + '; text-align: center; margin: 0 0 10pt; padding-bottom: 6pt; border-bottom: 2pt solid ' + accent + '; }',
    '.comp { font-size: 12pt; font-weight: bold; color: #fff; background: ' + accent + '; padding: 4pt 6pt; margin: 12pt 0 0; }',
    'table { width: 100%; border-collapse: collapse; page-break-inside: avoid; }',
    'th, td { border: 0.75pt solid #c0c7d4; padding: 4pt 3pt; text-align: center; }',
    'th { background: #d9e2f3; font-size: 8.5pt; }',
    'td.name { text-align: left; }',
    'td.place { font-weight: bold; font-size: 11pt; }',
    'td.best { font-weight: bold; }',
    'td.fail { color: #999; }',
    'tr.p1 td { background: #fff2cc; } tr.p2 td { background: #eeeeee; } tr.p3 td { background: #fce5cd; }',
    '.foot { margin-top: 10pt; font-size: 8pt; color: #888; text-align: right; }',
  ].join('\n');

  const pages = sections.map(function (s, si) {
    const tables = s.comps.map(function (c) {
      const ci = c.index;
      let head = '<tr><th>Miejsce</th><th>Zawodnik</th><th>Waga [kg]</th>';
      for (let i = 1; i <= c.attempts; i++) head += '<th>Próba ' + i + '</th>';
      head += '<th>Najlepsza [kg]</th><th>Wynik</th></tr>';

      const sorted = s.rows.slice().sort(function (a, b) {
        return a.comps[ci].place - b.comps[ci].place || (Number(a.nr) || 0) - (Number(b.nr) || 0);
      });
      const body = sorted.map(function (a) {
        const cc = a.comps[ci];
        const raw = (a.raw && a.raw[ci]) || [];
        const p = cc.score === null ? 0 : cc.place; // kolor podium tylko z wynikiem
        let tr = '<tr' + (p >= 1 && p <= 3 ? ' class="p' + p + '"' : '') + '>' +
          '<td class="place">' + cc.place + '</td><td class="name">' + escapeHtml_(a.name) + '</td>' +
          '<td>' + fmtPl_(a.bw, 1, false) + '</td>';
        for (let i = 0; i < c.attempts; i++) {
          const at = raw[i];
          tr += '<td' + (at && at.kg !== null && !at.ok ? ' class="fail"' : '') + '>' + escapeHtml_(attemptLabel_(at)) + '</td>';
        }
        tr += '<td class="best">' + fmtPl_(cc.best, 1, false) + '</td><td>' + fmtPl_(cc.score, decimals, true) + '</td></tr>';
        return tr;
      }).join('');

      return '<div class="comp">' + escapeHtml_(c.name) + '</div><table><thead>' + head + '</thead><tbody>' + body + '</tbody></table>';
    }).join('');

    return '<div class="page' + (si === sections.length - 1 ? ' last' : '') + '">' +
      '<div class="event">' + escapeHtml_(title) + '</div>' +
      '<div class="cat">' + escapeHtml_(s.category) + '</div>' + tables +
      '<div class="foot">Wydruk: ' + escapeHtml_(stamp) + '</div></div>';
  }).join('\n');

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' + css + '</style></head><body>' + pages + '</body></html>';
}
