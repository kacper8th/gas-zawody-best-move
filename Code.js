/**
 * Logika zawodów: listy startowe, przeliczanie wyników, blokady.
 *
 * Punktacja:
 *  - wynik konkurencji = najcięższa próba / waga_zawodnika ^ współczynnik (zaokrąglony wg konfiguracji),
 *    próba spalona = 0,
 *  - w konkurencji lepszy wynik = lepsze miejsce; remis = to samo miejsce, kolejne miejsce przeskakuje (1, 1, 3),
 *  - punkty = suma miejsc ze wszystkich konkurencji, mniej punktów = wyżej,
 *  - miejsce w dogrywce: 1 = -0,5 pkt, 2 = -0,33 pkt, 3 = -0,25 pkt (ostrzeżenie o remisie pokazuje się dla miejsc 1-3).
 */

/** Układ listy startowej. Kolumny stałe: Nr, Imię, Nazwisko, Waga. */
const LIST = { titleRow: 1, groupRow: 2, headerRow: 3, firstRow: 4, maxRows: 200, nrCol: 1, weightCol: 4, fixedCols: 4 };
const META_KEY = 'zbm_lista';
const PROP_LOCK_CONFIG = 'zbm_lock_config';
const PROP_LOCK_ATHLETES = 'zbm_lock_athletes';
/** Bonus za miejsce w dogrywce. */
const TIEBREAK_BONUS = { 1: -0.5, 2: -0.33, 3: -0.25 };
const TIEBREAK_MAX_PLACE = 3;

// ───────────────────────────── Menu ─────────────────────────────

function buildZawodyMenu_(ui) {
  const st = getLockState_();
  ui.createMenu('🏋️ Zawody')
    .addItem('Generuj listy startowe', 'generateStartLists')
    .addItem('Przelicz wyniki', 'rebuildResults')
    .addItem('✅ Walidacja zawodów', 'validateCompetition')
    .addItem('🖨️ Drukuj regulamin (PDF)', 'printRules')
    .addSeparator()
    .addItem(st.config ? '🔓 Odblokuj konfigurację' : '🔒 Zablokuj konfigurację', 'toggleConfigLock')
    .addItem(st.athletes ? '🔓 Odblokuj dane zawodników' : '🔒 Zablokuj dane zawodników', 'toggleAthletesLock')
    .addToUi();
}

// ───────────────────────── Listy startowe ─────────────────────────

function generateStartLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cfg = readConfig_(ss);

  const created = [], existing = [], errors = [];
  cfg.categories.forEach(function (cat) {
    if (!(cat.coef >= 0 && cat.coef <= 1)) { errors.push(cat.name + ': brak lub błędny współczynnik (0-1)'); return; }
    if (!cat.comps.length) { errors.push(cat.name + ': nie wybrano żadnej konkurencji'); return; }
    const sh = findListSheet_(ss, cat.name);
    if (sh) { existing.push(cat.name + ' (zakładka „' + sh.getName() + '”)'); return; }
    buildListSheet_(ss, cat, cfg);
    created.push(cat.name);
  });

  rebuildResults();

  const msg = [];
  if (created.length) msg.push('Utworzono listy:\n• ' + created.join('\n• '));
  if (existing.length) msg.push('Lista już istnieje - aby utworzyć nową, najpierw usuń starą zakładkę:\n• ' + existing.join('\n• '));
  if (errors.length) msg.push('Pominięto (popraw konfigurację):\n• ' + errors.join('\n• '));
  if (!msg.length) msg.push('Brak kategorii w konfiguracji.');
  ui.alert('Listy startowe', msg.join('\n\n'), ui.ButtonSet.OK);
}

function buildListSheet_(ss, cat, cfg) {
  const results = ss.getSheetByName(SHEET_RESULTS);
  const index = results ? results.getIndex() - 1 : ss.getSheets().length;
  const sh = ss.insertSheet(listSheetName_(cat.name), index);

  const headers = ['Nr', 'Imię', 'Nazwisko', 'Waga [kg]'];
  const groups = ['', 'Zawodnik', '', ''];
  const comps = [];
  cat.comps.forEach(function (c) {
    comps.push({
      key: c.key, name: c.name, short: c.short, col: headers.length + 1, attempts: c.attempts,
      refBodyweight: c.refBodyweight, refWeight: c.refWeight,
    });
    // Każda próba = 2 kolumny: ciężar [kg] i checkbox „Zal.” (zaliczona).
    for (let i = 1; i <= c.attempts; i++) {
      headers.push('Próba ' + i, ''); // nagłówek scalony nad [kg] i „Zal.”
      groups.push(i === 1 ? c.name + ' [kg]' : '', '');
    }
  });
  headers.push('Dogrywka - miejsce');
  groups.push('Dogrywka');
  const tieCol = headers.length;
  const nCols = headers.length;
  const lastRow = LIST.firstRow + LIST.maxRows - 1;

  // Przycięcie siatki do potrzebnego rozmiaru.
  if (sh.getMaxColumns() > nCols) sh.deleteColumns(nCols + 1, sh.getMaxColumns() - nCols);
  if (sh.getMaxColumns() < nCols) sh.insertColumnsAfter(sh.getMaxColumns(), nCols - sh.getMaxColumns());
  if (sh.getMaxRows() > lastRow) sh.deleteRows(lastRow + 1, sh.getMaxRows() - lastRow);

  sh.getRange(LIST.titleRow, 1)
    .setValue('Lista startowa - ' + cat.name + (cfg.title ? ' · ' + cfg.title : ''))
    .setFontSize(14).setFontWeight('bold');

  sh.getRange(LIST.groupRow, 1, 1, nCols).setValues([groups])
    .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#d9e2f3');
  sh.getRange(LIST.groupRow, 2, 1, LIST.fixedCols - 1).merge();
  comps.forEach(function (c) {
    sh.getRange(LIST.groupRow, c.col, 1, 2 * c.attempts).merge();
    // Nazwa konkurencji z konfiguracji - zmiana nazwy od razu widoczna na liście.
    const idx = COMPETITIONS.map(function (k) { return k.key; }).indexOf(c.key);
    sh.getRange(LIST.groupRow, c.col).setFormula('=' + compCellA1_(idx, 'name', true) + '&" [kg]"');
  });

  sh.getRange(LIST.headerRow, 1, 1, nCols).setValues([headers])
    .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#eef3fb').setWrap(true);

  // Numeracja od 1 - pojawia się, gdy wpisane jest imię lub nazwisko.
  const r1 = LIST.firstRow;
  setNumberingFormula_(sh);

  const kgRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false)
    .setHelpText('Ciężar w kg (dokładność 0,1). Próba zaliczona = zaznacz „Zal.” obok.').build();
  const weightRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0).setAllowInvalid(false)
    .setHelpText('Waga zawodnika w kg (dokładność 0,1).').build();

  sh.getRange(r1, LIST.nrCol, LIST.maxRows, 1).setHorizontalAlignment('center').setFontColor('#666666').setBackground('#f3f3f3');
  sh.getRange(r1, LIST.weightCol, LIST.maxRows, 1).setDataValidation(weightRule).setNumberFormat('0.0').setHorizontalAlignment('center');
  comps.forEach(function (c) {
    for (let i = 0; i < c.attempts; i++) {
      sh.getRange(r1, c.col + 2 * i, LIST.maxRows, 1).setDataValidation(kgRule).setNumberFormat('0.0').setHorizontalAlignment('center');
      sh.getRange(r1, c.col + 2 * i + 1, LIST.maxRows, 1).insertCheckboxes();
    }
  });
  sh.getRange(r1, tieCol, LIST.maxRows, 1).setHorizontalAlignment('center')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.keys(TIEBREAK_BONUS), true).setAllowInvalid(false)
      .setHelpText('Miejsce w dogrywce: 1 = -0,5 pkt, 2 = -0,33 pkt, 3 = -0,25 pkt. Puste = bez dogrywki.').build());

  sh.getRange(LIST.groupRow, 1, lastRow - LIST.groupRow + 1, nCols)
    .setBorder(true, true, true, true, true, true, '#c0c7d4', SpreadsheetApp.BorderStyle.SOLID);

  // Wizualny podział: para [ciężar | Zal.] w ramce, konkurencje i dogrywka oddzielone grubą linią.
  const H = LIST.headerRow, nBody = lastRow - LIST.headerRow + 1;
  comps.forEach(function (c) {
    for (let i = 0; i < c.attempts; i++) {
      const col = c.col + 2 * i;
      sh.getRange(H, col, 1, 2).merge();
      sh.getRange(H, col, nBody, 2).setBorder(null, true, null, true, null, null, '#8896ab', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
    sh.getRange(LIST.groupRow, c.col, nBody + 1, 2 * c.attempts)
      .setBorder(null, true, null, true, null, null, '#1f4e79', SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  sh.getRange(LIST.groupRow, tieCol, nBody + 1, 1)
    .setBorder(null, true, null, true, null, null, '#1f4e79', SpreadsheetApp.BorderStyle.SOLID_THICK);

  sh.setColumnWidth(1, 45);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(4, 85);
  for (let c = LIST.fixedCols + 1; c < tieCol; c++) sh.setColumnWidth(c, (c - LIST.fixedCols) % 2 ? 65 : 38);
  sh.setColumnWidth(tieCol, 95);
  sh.setFrozenRows(LIST.headerRow);
  sh.setFrozenColumns(LIST.fixedCols);

  sh.addDeveloperMetadata(META_KEY, JSON.stringify({
    layout: 2, category: cat.name, coef: cat.coef, comps: comps, tieCol: tieCol,
  }));

  // Nr jest liczony - chroniony zawsze (dla właściciela pilnuje tego onEdit).
  hardProtect_(sh.getRange(r1, LIST.nrCol, LIST.maxRows, 1).protect().setDescription('zbm:nr'));
  if (getLockState_().athletes) protectAthletes_(sh);
  return sh;
}

/** Numeracja od 1 - pojawia się, gdy wpisane jest imię lub nazwisko. */
function setNumberingFormula_(sh) {
  const r1 = LIST.firstRow;
  const lastRow = LIST.firstRow + LIST.maxRows - 1;
  sh.getRange(r1, LIST.nrCol, LIST.maxRows, 1).clearContent();
  sh.getRange(r1, LIST.nrCol).setFormula(
    '=ARRAYFORMULA(IF((B' + r1 + ':B' + lastRow + '&C' + r1 + ':C' + lastRow + ')<>"";ROW(B' + r1 + ':B' + lastRow + ')-' + (r1 - 1) + ';""))');
}

function listSheetName_(category) {
  return (LIST_PREFIX + category).replace(/[\[\]\*\?\/\\:]/g, '-').slice(0, 100);
}

function getListMeta_(sh) {
  const md = sh.getDeveloperMetadata().filter(function (d) { return d.getKey() === META_KEY; })[0];
  if (!md) return null;
  try { return JSON.parse(md.getValue()); } catch (e) { return null; }
}

function getListSheets_(ss) {
  return ss.getSheets()
    .map(function (sh) { return { sheet: sh, meta: getListMeta_(sh) }; })
    .filter(function (x) { return x.meta; });
}

function findListSheet_(ss, category) {
  const byMeta = getListSheets_(ss).filter(function (x) { return x.meta.category === category; })[0];
  return byMeta ? byMeta.sheet : ss.getSheetByName(listSheetName_(category));
}

// ───────────────────────────── Wyniki ─────────────────────────────

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() === SHEET_RESULTS) return;
  const isList = !!getListMeta_(sh);
  const isConfig = sh.getName() === SHEET_CONFIG;
  if (!isConfig && !isList) return;

  // Blokady: ochrona Google nie działa na właściciela, więc ręczne zmiany w zablokowanych obszarach są cofane.
  const st = getLockState_();
  if (isConfig && st.config) {
    if (restoreFromSnapshot_(sh, SNAP_CONFIG, e.range)) {
      sh.getParent().toast('Konfiguracja jest zablokowana - zmiana cofnięta.', '🔒 Blokada');
    }
    return;
  }
  if (isList && st.athletes && restoreFromSnapshot_(sh, athletesSnapKey_(sh), e.range)) {
    sh.getParent().toast('Dane zawodników są zablokowane - zmiana cofnięta.', '🔒 Blokada');
  }

  // Ręczna zmiana w kolumnie Nr - przywróć automatyczną numerację.
  if (isList && e.range.getColumn() <= LIST.nrCol && e.range.getLastRow() >= LIST.firstRow) {
    setNumberingFormula_(sh);
    sh.getParent().toast('Nr jest liczony automatycznie - nie można go zmieniać.', '🏋️ Zawody');
  }
  try {
    rebuildResults();
  } catch (err) {
    console.error(err);
  }
}

function rebuildResults() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return;
  try {
    rebuildResultsUnlocked_();
  } finally {
    lock.releaseLock();
  }
}

function rebuildResultsUnlocked_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = readConfig_(ss);
  const sh = ss.getSheetByName(SHEET_RESULTS) || createResultsSheet_(ss);

  const order = cfg.categories.map(function (c) { return c.name; });
  const rank = function (x) { const i = order.indexOf(x.meta.category); return i < 0 ? 999 : i; };
  const lists = getListSheets_(ss).sort(function (a, b) { return rank(a) - rank(b); });

  const scoreFmt = cfg.decimals > 0 ? '0.' + '0'.repeat(cfg.decimals) : '0';
  const rows = [];
  const fmt = []; // { row, kind, width, scoreCols }

  rows.push([cfg.title ? 'Wyniki - ' + cfg.title : 'Wyniki']);
  fmt.push({ row: rows.length, kind: 'title' });
  rows.push(['Przeliczono: ' + Utilities.formatDate(new Date(), 'Europe/Warsaw', 'yyyy-MM-dd HH:mm:ss')]);
  fmt.push({ row: rows.length, kind: 'info' });

  if (!lists.length) {
    rows.push(['']);
    rows.push(['Brak list startowych - 🏋️ Zawody → Generuj listy startowe.']);
    fmt.push({ row: rows.length, kind: 'info' });
  }

  lists.forEach(function (x) {
    const meta = x.meta;
    const res = categoryResult_(cfg, x.sheet, meta);

    rows.push(['']);
    rows.push([meta.category]);
    fmt.push({ row: rows.length, kind: 'section' });

    // Wiersz grup (nazwa konkurencji nad jej 3 kolumnami, „Podsumowanie”) + nagłówki.
    const header = ['Miejsce', 'Zawodnik', 'Waga [kg]'];
    const groupsRow = ['', '', ''];
    const blocks = []; // [kolumna startowa, szerokość]
    const scoreCols = [];
    meta.comps.forEach(function (c) {
      const live = cfg.competitions.filter(function (k) { return k.key === c.key; })[0];
      blocks.push([header.length + 1, 3]);
      groupsRow.push((live || c).name, '', '');
      header.push('Najlepsza [kg]', 'Wynik', 'Miejsce');
      scoreCols.push(header.length - 1);
    });
    blocks.push([header.length + 1, 3]);
    groupsRow.push('Podsumowanie', '', '');
    header.push('Suma miejsc', 'Dogrywka', 'Punkty', 'Uwagi');
    rows.push(groupsRow);
    const groupRowNo = rows.length;
    fmt.push({ row: groupRowNo, kind: 'groups', blocks: blocks });
    rows.push(header);
    fmt.push({ row: rows.length, kind: 'header', width: header.length });

    if (!res.rows.length) {
      rows.push(['', '(brak zawodników)']);
      fmt.push({ row: rows.length, kind: 'info' });
    }

    const firstBody = rows.length + 1;
    res.rows.forEach(function (a) {
      const r = [a.place, a.name, a.bw === null ? '' : a.bw];
      a.comps.forEach(function (c) {
        r.push(c.best === null ? '' : c.best, c.score === null ? '' : c.score, c.place);
      });
      r.push(a.sum, a.tiebreak || '', a.points, a.notes.join('; '));
      rows.push(r);
    });
    if (res.rows.length) {
      fmt.push({ row: firstBody, kind: 'body', count: res.rows.length, width: header.length, scoreCols: scoreCols });
    }

    fmt.push({ row: groupRowNo, kind: 'separators', count: rows.length - groupRowNo + 1, blocks: blocks, width: header.length });

    res.unresolved.forEach(function (u) {
      rows.push(['', '⚠ Remis na miejscu ' + u.place + ': ' + u.names.join(', ') + ' - zaznacz zwycięzcę dogrywki na liście startowej.']);
      fmt.push({ row: rows.length, kind: 'warn' });
    });
  });

  // Zapis jednym setValues.
  const width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
  const grid = rows.map(function (r) { return r.concat(new Array(width - r.length).fill('')); });

  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  if (sh.getMaxRows() < grid.length) sh.insertRowsAfter(sh.getMaxRows(), grid.length - sh.getMaxRows());
  sh.getRange(1, 1, grid.length, width).setValues(grid).setVerticalAlignment('middle');

  fmt.forEach(function (f) {
    if (f.kind === 'title') sh.getRange(f.row, 1).setFontSize(16).setFontWeight('bold');
    if (f.kind === 'info') sh.getRange(f.row, 1, 1, 2).setFontColor('#666666').setFontStyle('italic');
    if (f.kind === 'section') sh.getRange(f.row, 1).setFontSize(13).setFontWeight('bold');
    if (f.kind === 'warn') sh.getRange(f.row, 2).setFontColor('#c00000').setFontWeight('bold');
    if (f.kind === 'groups') {
      f.blocks.forEach(function (b) {
        sh.getRange(f.row, b[0], 1, b[1]).merge()
          .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#1f4e79').setFontColor('#ffffff');
      });
    }
    if (f.kind === 'separators') {
      f.blocks.forEach(function (b) {
        sh.getRange(f.row, b[0], f.count, b[1])
          .setBorder(null, true, null, true, null, null, '#1f4e79', SpreadsheetApp.BorderStyle.SOLID_THICK);
      });
    }
    if (f.kind === 'header') {
      sh.getRange(f.row, 1, 1, f.width).setFontWeight('bold').setBackground('#d9e2f3').setWrap(true).setHorizontalAlignment('center');
    }
    if (f.kind === 'body') {
      const body = sh.getRange(f.row, 1, f.count, f.width);
      body.setHorizontalAlignment('center');
      sh.getRange(f.row, 2, f.count, 1).setHorizontalAlignment('left');
      sh.getRange(f.row, f.width, f.count, 1).setHorizontalAlignment('left').setFontColor('#666666');
      sh.getRange(f.row, 1, f.count, 1).setFontWeight('bold');
      sh.getRange(f.row, 3, f.count, 1).setNumberFormat('0.0');
      f.scoreCols.forEach(function (col) {
        sh.getRange(f.row, col - 1, f.count, 1).setNumberFormat('0.0');
        sh.getRange(f.row, col, f.count, 1).setNumberFormat(scoreFmt);
      });
      sh.getRange(f.row, f.width - 1, f.count, 1).setFontWeight('bold');
      body.setBorder(true, true, true, true, true, true, '#c0c7d4', SpreadsheetApp.BorderStyle.SOLID);
      // Podium.
      const podium = ['#fff2cc', '#eeeeee', '#fce5cd'];
      for (let i = 0; i < f.count; i++) {
        const place = grid[f.row - 1 + i][0];
        if (place >= 1 && place <= 3) sh.getRange(f.row + i, 1, 1, f.width).setBackground(podium[place - 1]);
      }
    }
  });
}

/** Czyta zawodników z listy startowej (pomija puste wiersze). */
/** Wynik kategorii: współczynnik i referencje z bieżącej konfiguracji (gdy kategorii już nie ma - z chwili generowania listy). */
function categoryResult_(cfg, sheet, meta) {
  const cfgCat = cfg.categories.filter(function (c) { return c.name === meta.category; })[0];
  const coef = cfgCat && cfgCat.coef >= 0 && cfgCat.coef <= 1 ? cfgCat.coef : meta.coef;
  const refs = meta.comps.map(function (c) {
    const live = cfgCat && cfgCat.comps.filter(function (k) { return k.key === c.key; })[0];
    return live || c;
  });
  return computeRanking_(readAthletes_(sheet, meta), refs, coef, cfg.decimals);
}

/** Próby konkurencji z wiersza listy: [{ kg, ok, col }]. Stary układ (bez „Zal.”) = każda wpisana próba zaliczona. */
function readAttempts_(row, c, meta) {
  const out = [];
  for (let i = 0; i < c.attempts; i++) {
    const col = meta.layout === 2 ? c.col + 2 * i : c.col + i;
    out.push({ kg: toNum_(row[col - 1]), ok: meta.layout === 2 ? row[col] === true : true, col: col });
  }
  return out;
}

/** Czyta zawodników z listy startowej (pomija puste wiersze). Do wyniku liczą się tylko próby zaliczone. */
function readAthletes_(sheet, meta) {
  const values = sheet.getRange(LIST.firstRow, 1, LIST.maxRows, meta.tieCol).getValues();
  const out = [];
  values.forEach(function (r) {
    const first = String(r[1]).trim();
    const last = String(r[2]).trim();
    if (!first && !last) return;
    out.push({
      nr: r[0],
      name: (first + ' ' + last).trim(),
      bw: toNum_(r[LIST.weightCol - 1]),
      tiePlace: toNum_(r[meta.tieCol - 1]),
      attempts: meta.comps.map(function (c) {
        return readAttempts_(r, c, meta)
          .filter(function (a) { return a.ok && a.kg !== null; })
          .map(function (a) { return a.kg; });
      }),
    });
  });
  return out;
}

// ───────────────────────────── Walidacja ─────────────────────────────

/**
 * Sprawdza listy startowe: brak wagi, puste próby, próba lżejsza od poprzedniej, zaliczona bez ciężaru,
 * nierozstrzygnięte remisy na podium. Błędne komórki podświetla na czerwono.
 */
function validateCompetition() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cfg = readConfig_(ss);
  const lists = getListSheets_(ss);
  const errors = [], warnings = [];
  const BAD = '#f4cccc';

  if (!lists.length) {
    ui.alert('Walidacja zawodów', 'Brak list startowych.', ui.ButtonSet.OK);
    return;
  }

  lists.forEach(function (x) {
    const sh = x.sheet, meta = x.meta;
    const res = validateList_(sh.getRange(LIST.firstRow, 1, LIST.maxRows, meta.tieCol).getValues(), meta, cfg);

    // Czyszczenie starych podświetleń (waga + próby) i nowe podświetlenia.
    sh.getRange(LIST.firstRow, LIST.weightCol, LIST.maxRows, 1).setBackground(null);
    meta.comps.forEach(function (c) {
      sh.getRange(LIST.firstRow, c.col, LIST.maxRows, meta.layout === 2 ? 2 * c.attempts : c.attempts).setBackground(null);
    });
    res.cells.forEach(function (rc) { sh.getRange(rc[0], rc[1]).setBackground(BAD); });

    res.errors.forEach(function (m) { errors.push(meta.category + ': ' + m); });
    categoryResult_(cfg, sh, meta).unresolved.forEach(function (u) {
      warnings.push(meta.category + ': nierozstrzygnięty remis na miejscu ' + u.place + ' (' + u.names.join(', ') + ') - uzupełnij dogrywkę');
    });
  });

  if (!errors.length && !warnings.length) {
    ui.alert('Walidacja zawodów', '✅ Wszystko poprawne - brak błędów.', ui.ButtonSet.OK);
    return;
  }
  const MAX = 30;
  const lines = [];
  if (errors.length) {
    lines.push('❌ Błędy (' + errors.length + '):');
    errors.slice(0, MAX).forEach(function (m) { lines.push('• ' + m); });
    if (errors.length > MAX) lines.push('… i ' + (errors.length - MAX) + ' więcej (podświetlone na czerwono na listach).');
  }
  if (warnings.length) {
    if (lines.length) lines.push('');
    lines.push('⚠ Remisy na podium (' + warnings.length + '):');
    warnings.forEach(function (m) { lines.push('• ' + m); });
  }
  ui.alert('Walidacja zawodów', lines.join('\n'), ui.ButtonSet.OK);
}

/** Czysta walidacja wierszy listy. Zwraca { errors: [tekst], cells: [[wiersz, kolumna]] }. */
function validateList_(values, meta, cfg) {
  const errors = [], cells = [];
  values.forEach(function (r, k) {
    const first = String(r[1]).trim();
    const last = String(r[2]).trim();
    if (!first && !last) return;
    const row = LIST.firstRow + k;
    const who = 'nr ' + r[0] + ' ' + (first + ' ' + last).trim();

    if (toNum_(r[LIST.weightCol - 1]) === null) {
      errors.push(who + ' - brak wagi');
      cells.push([row, LIST.weightCol]);
    }
    meta.comps.forEach(function (c) {
      const live = cfg && cfg.competitions ? cfg.competitions.filter(function (q) { return q.key === c.key; })[0] : null;
      const short = (live || c).short;
      let prev = null, prevNo = 0;
      readAttempts_(r, c, meta).forEach(function (a, i) {
        if (a.kg === null) {
          errors.push(who + ' - ' + short + ': brak próby ' + (i + 1));
          cells.push([row, a.col]);
          return;
        }
        if (prev !== null && a.kg < prev) {
          errors.push(who + ' - ' + short + ': próba ' + (i + 1) + ' (' + a.kg + ' kg) lżejsza niż próba ' + prevNo + ' (' + prev + ' kg)');
          cells.push([row, a.col]);
        }
        prev = a.kg;
        prevNo = i + 1;
      });
    });
  });
  return { errors: errors, cells: cells };
}

/**
 * Liczy wyniki kategorii. Czysta funkcja (bez API arkusza) - łatwa do przetestowania.
 * athletes: [{ nr, name, bw, tiePlace, attempts: [[kg, ...] per konkurencja] }]
 * refs: [{ refBodyweight, refWeight } per konkurencja]
 *
 * wynik = ciężar / waga^wsp * waga_ref^wsp / ciężar_ref
 * (zawodnik referencyjny = 1; referencja tylko skaluje wynik, kolejności nie zmienia).
 * Brak referencji = wynik bez skalowania.
 */
function computeRanking_(athletes, refs, coef, decimals) {
  const nComps = refs.length;
  const factors = refs.map(function (r) {
    return r && r.refBodyweight > 0 && r.refWeight > 0 ? Math.pow(r.refBodyweight, coef) / r.refWeight : 1;
  });
  const rows = athletes.map(function (a) {
    const notes = [];
    if (a.bw === null) notes.push('brak wagi');
    const comps = [];
    for (let i = 0; i < nComps; i++) {
      const att = a.attempts[i] || [];
      const best = att.length ? Math.max.apply(null, att) : null;
      const score = best !== null && a.bw ? round_(best / Math.pow(a.bw, coef) * factors[i], decimals) : null;
      comps.push({ best: best, score: score, place: null });
    }
    return { nr: a.nr, name: a.name, bw: a.bw, tiePlace: a.tiePlace, comps: comps, notes: notes };
  });

  // Miejsca w konkurencjach: większy wynik = lepiej; brak wyniku = na końcu.
  for (let i = 0; i < nComps; i++) {
    const places = competitionRank_(rows.map(function (r) {
      return r.comps[i].score === null ? -Infinity : r.comps[i].score;
    }), true);
    rows.forEach(function (r, k) { r.comps[i].place = places[k]; });
  }

  rows.forEach(function (r) {
    r.sum = r.comps.reduce(function (s, c) { return s + c.place; }, 0);
  });

  // Dogrywka: miejsce 1/2/3 daje -0,5 / -0,33 / -0,25 pkt.
  rows.forEach(function (r) {
    r.tiebreak = TIEBREAK_BONUS[r.tiePlace] || 0;
    r.points = round_(r.sum + r.tiebreak, 2);
  });

  const finalPlaces = competitionRank_(rows.map(function (r) { return r.points; }), false);
  rows.forEach(function (r, k) { r.place = finalPlaces[k]; });

  // Nierozstrzygnięte remisy na podium.
  const unresolved = [];
  const seen = {};
  rows.forEach(function (r) {
    if (r.place > TIEBREAK_MAX_PLACE || seen[r.place]) return;
    const group = rows.filter(function (o) {
      return o.place === r.place && o.comps.some(function (c) { return c.score !== null; });
    });
    if (group.length > 1) {
      seen[r.place] = true;
      unresolved.push({ place: r.place, names: group.map(function (o) { return o.name; }) });
    }
  });

  rows.sort(function (a, b) { return a.place - b.place || (Number(a.nr) || 0) - (Number(b.nr) || 0); });
  return { rows: rows, unresolved: unresolved };
}

/** Miejsca „sportowe”: remis = to samo miejsce, kolejne przeskakuje (1, 1, 3). */
function competitionRank_(values, higherIsBetter) {
  return values.map(function (v) {
    return 1 + values.filter(function (w) { return higherIsBetter ? w > v : w < v; }).length;
  });
}

function round_(x, decimals) {
  const f = Math.pow(10, decimals);
  return Math.round((x + Number.EPSILON) * f) / f;
}

function toNum_(v) {
  if (v === '' || v === null || v === undefined || v === true || v === false) return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;
  const n = Number(String(v).replace(',', '.').trim());
  return isNaN(n) ? null : n;
}

// ───────────────────────────── Blokady ─────────────────────────────
// Ochrona Google nie blokuje właściciela arkusza ani tego, kto ją założył - blokuje pozostałych edytorów.

function getLockState_() {
  try {
    const p = PropertiesService.getDocumentProperties();
    return { config: p.getProperty(PROP_LOCK_CONFIG) === '1', athletes: p.getProperty(PROP_LOCK_ATHLETES) === '1' };
  } catch (e) {
    return { config: false, athletes: false };
  }
}

function toggleConfigLock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) throw new Error('Brak zakładki „' + SHEET_CONFIG + '”.');
  const props = PropertiesService.getDocumentProperties();
  const locked = getLockState_().config;

  if (locked) {
    removeProtections_(sh, SpreadsheetApp.ProtectionType.SHEET, 'zbm:config');
    props.deleteProperty(SNAP_CONFIG);
    props.setProperty(PROP_LOCK_CONFIG, '0');
    sh.getRange(CFG.lockConfigCell).setValue(lockLabel_(false));
  } else {
    sh.getRange(CFG.lockConfigCell).setValue(lockLabel_(true));
    SpreadsheetApp.flush();
    saveSnapshot_(sh, SNAP_CONFIG, configArea_(), false);
    hardProtect_(sh.protect().setDescription('zbm:config'));
    props.setProperty(PROP_LOCK_CONFIG, '1');
  }
  buildMenus_();
  ss.toast(locked ? 'Konfiguracja odblokowana.' : 'Konfiguracja zablokowana.', '🏋️ Zawody');
}

/** Blokuje imię, nazwisko i wagę na wszystkich listach - nie da się dodać ani zmienić zawodnika. */
function toggleAthletesLock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const locked = getLockState_().athletes;

  getListSheets_(ss).forEach(function (x) {
    if (locked) {
      removeProtections_(x.sheet, SpreadsheetApp.ProtectionType.RANGE, 'zbm:zawodnicy');
      props.deleteProperty(athletesSnapKey_(x.sheet));
    } else {
      protectAthletes_(x.sheet);
    }
  });
  props.setProperty(PROP_LOCK_ATHLETES, locked ? '0' : '1');

  const cfgSheet = ss.getSheetByName(SHEET_CONFIG);
  if (cfgSheet) {
    try {
      cfgSheet.getRange(CFG.lockAthletesCell).setValue(lockLabel_(!locked));
      // Konfiguracja zablokowana - odśwież jej zapamiętany stan, żeby nie cofać nowego statusu.
      if (getLockState_().config) { SpreadsheetApp.flush(); saveSnapshot_(cfgSheet, SNAP_CONFIG, configArea_(), false); }
    } catch (e) { /* konfiguracja zablokowana dla tego użytkownika */ }
  }
  buildMenus_();
  ss.toast(locked ? 'Dane zawodników odblokowane.' : 'Dane zawodników zablokowane.', '🏋️ Zawody');
}

function protectAthletes_(sheet) {
  removeProtections_(sheet, SpreadsheetApp.ProtectionType.RANGE, 'zbm:zawodnicy');
  const a = athletesArea_();
  const p = sheet.getRange(a.row, a.col, a.rows, a.cols).protect().setDescription('zbm:zawodnicy');
  hardProtect_(p);
  saveSnapshot_(sheet, athletesSnapKey_(sheet), a, true);
}

// ── Zapamiętany stan zablokowanych obszarów (do cofania zmian w onEdit) ──

const SNAP_CONFIG = 'zbm_snap_config';

function athletesSnapKey_(sheet) {
  return 'zbm_snap_ath_' + sheet.getSheetId();
}

/** Imię, Nazwisko, Waga na liście startowej. */
function athletesArea_() {
  return { row: LIST.firstRow, col: 2, rows: LIST.maxRows, cols: LIST.fixedCols - 1 };
}

/** Cała część zakładki Konfiguracja z ustawieniami i tabelami. */
function configArea_() {
  const n = COMPETITIONS.length;
  return { row: 1, col: 1, rows: CFG.refFirstRow + CFG.maxRows, cols: Math.max(2 + 2 * n, 1 + 2 * n, CFG.compShortCol) };
}

/** Zapisuje wartości (formuły jako tekst formuły) obszaru do właściwości dokumentu. */
function saveSnapshot_(sheet, key, area, trimEmptyRows) {
  const range = sheet.getRange(area.row, area.col, area.rows, area.cols);
  const formulas = range.getFormulas();
  const cells = range.getValues().map(function (row, i) {
    return row.map(function (v, j) { return formulas[i][j] || v; });
  });
  if (trimEmptyRows) {
    while (cells.length && cells[cells.length - 1].every(function (v) { return v === ''; })) cells.pop();
  }
  const snap = { row: area.row, col: area.col, rows: area.rows, cols: area.cols, cells: cells };
  PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify(snap));
}

/** Przywraca zapamiętany stan w części edytowanego zakresu, która należy do obszaru. Zwraca true, jeśli coś cofnięto. */
function restoreFromSnapshot_(sheet, key, edited) {
  const raw = PropertiesService.getDocumentProperties().getProperty(key);
  if (!raw) return false;
  const snap = JSON.parse(raw);
  const box = intersect_(
    { r1: edited.getRow(), r2: edited.getLastRow(), c1: edited.getColumn(), c2: edited.getLastColumn() },
    { r1: snap.row, r2: snap.row + snap.rows - 1, c1: snap.col, c2: snap.col + snap.cols - 1 });
  if (!box) return false;
  sheet.getRange(box.r1, box.c1, box.r2 - box.r1 + 1, box.c2 - box.c1 + 1).setValues(snapshotCells_(snap, box));
  return true;
}

/** Część wspólna dwóch prostokątów {r1, r2, c1, c2} albo null. */
function intersect_(a, b) {
  const box = { r1: Math.max(a.r1, b.r1), r2: Math.min(a.r2, b.r2), c1: Math.max(a.c1, b.c1), c2: Math.min(a.c2, b.c2) };
  return box.r1 <= box.r2 && box.c1 <= box.c2 ? box : null;
}

/** Wartości ze snapshotu dla prostokąta (wiersze poza zapamiętanymi = puste). */
function snapshotCells_(snap, box) {
  const out = [];
  for (let r = box.r1; r <= box.r2; r++) {
    const src = snap.cells[r - snap.row] || [];
    const row = [];
    for (let c = box.c1; c <= box.c2; c++) {
      const v = src[c - snap.col];
      row.push(v === undefined || v === null ? '' : v);
    }
    out.push(row);
  }
  return out;
}

function hardProtect_(protection) {
  const me = Session.getEffectiveUser();
  protection.addEditor(me);
  const others = protection.getEditors().filter(function (u) { return u.getEmail() !== me.getEmail(); });
  if (others.length) protection.removeEditors(others);
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
  return protection;
}

function removeProtections_(sheet, type, description) {
  sheet.getProtections(type)
    .filter(function (p) { return p.getDescription() === description; })
    .forEach(function (p) { p.remove(); });
}
