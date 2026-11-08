/**
 * Logika zawodów: listy startowe, przeliczanie wyników, blokady.
 *
 * Punktacja:
 *  - wynik konkurencji = najcięższa próba / waga_zawodnika ^ współczynnik (zaokrąglony wg konfiguracji),
 *    próba spalona = 0,
 *  - w konkurencji lepszy wynik = lepsze miejsce; remis = to samo miejsce, kolejne miejsce przeskakuje (1, 1, 3),
 *  - punkty = suma miejsc ze wszystkich konkurencji, mniej punktów = wyżej,
 *  - zaznaczona wygrana dogrywka zawsze daje -0,5 pkt (ostrzeżenie o remisie pokazuje się dla miejsc 1-3).
 */

/** Układ listy startowej. Kolumny stałe: Nr, Imię, Nazwisko, Waga. */
const LIST = { titleRow: 1, groupRow: 2, headerRow: 3, firstRow: 4, maxRows: 200, nrCol: 1, weightCol: 4, fixedCols: 4 };
const META_KEY = 'zbm_lista';
const PROP_LOCK_CONFIG = 'zbm_lock_config';
const PROP_LOCK_ATHLETES = 'zbm_lock_athletes';
const TIEBREAK_BONUS = -0.5;
const TIEBREAK_MAX_PLACE = 3;

// ───────────────────────────── Menu ─────────────────────────────

function buildZawodyMenu_(ui) {
  const st = getLockState_();
  ui.createMenu('🏋️ Zawody')
    .addItem('Generuj listy startowe', 'generateStartLists')
    .addItem('Przelicz wyniki', 'rebuildResults')
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
    for (let i = 1; i <= c.attempts; i++) {
      headers.push('Próba ' + i);
      groups.push(i === 1 ? c.name + ' [kg]' : '');
    }
  });
  headers.push('Wygrana dogrywka');
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
    if (c.attempts > 1) sh.getRange(LIST.groupRow, c.col, 1, c.attempts).merge();
  });

  sh.getRange(LIST.headerRow, 1, 1, nCols).setValues([headers])
    .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#eef3fb').setWrap(true);

  // Numeracja od 1 - pojawia się, gdy wpisane jest imię lub nazwisko.
  const r1 = LIST.firstRow;
  setNumberingFormula_(sh);

  const kgRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false)
    .setHelpText('Ciężar w kg (dokładność 0,1). Spalona próba = 0.').build();
  const weightRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0).setAllowInvalid(false)
    .setHelpText('Waga zawodnika w kg (dokładność 0,1).').build();

  sh.getRange(r1, LIST.nrCol, LIST.maxRows, 1).setHorizontalAlignment('center').setFontColor('#666666').setBackground('#f3f3f3');
  sh.getRange(r1, LIST.weightCol, LIST.maxRows, 1).setDataValidation(weightRule).setNumberFormat('0.0').setHorizontalAlignment('center');
  comps.forEach(function (c) {
    sh.getRange(r1, c.col, LIST.maxRows, c.attempts).setDataValidation(kgRule).setNumberFormat('0.0').setHorizontalAlignment('center');
  });
  sh.getRange(r1, tieCol, LIST.maxRows, 1).insertCheckboxes();

  sh.getRange(LIST.groupRow, 1, lastRow - LIST.groupRow + 1, nCols)
    .setBorder(true, true, true, true, true, true, '#c0c7d4', SpreadsheetApp.BorderStyle.SOLID);

  sh.setColumnWidth(1, 45);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(4, 85);
  for (let c = LIST.fixedCols + 1; c < tieCol; c++) sh.setColumnWidth(c, 75);
  sh.setColumnWidth(tieCol, 95);
  sh.setFrozenRows(LIST.headerRow);
  sh.setFrozenColumns(LIST.fixedCols);

  sh.addDeveloperMetadata(META_KEY, JSON.stringify({
    category: cat.name, coef: cat.coef, comps: comps, tieCol: tieCol,
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
  if (sh.getName() !== SHEET_CONFIG && !isList) return;
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
    const cfgCat = cfg.categories.filter(function (c) { return c.name === meta.category; })[0];
    const coef = cfgCat && cfgCat.coef >= 0 && cfgCat.coef <= 1 ? cfgCat.coef : meta.coef;

    const athletes = readAthletes_(x.sheet, meta);
    const res = computeRanking_(athletes, meta.comps.length, coef, cfg.decimals);

    rows.push(['']);
    rows.push([meta.category + '   (współczynnik ' + coef + ')']);
    fmt.push({ row: rows.length, kind: 'section' });

    const header = ['Miejsce', 'Zawodnik', 'Nr', 'Waga [kg]'];
    const scoreCols = [];
    meta.comps.forEach(function (c) {
      header.push(c.short + ' - najlepsza [kg]', c.short + ' - wynik', c.short + ' - miejsce');
      scoreCols.push(header.length - 1);
    });
    header.push('Suma miejsc', 'Dogrywka', 'Punkty', 'Uwagi');
    rows.push(header);
    fmt.push({ row: rows.length, kind: 'header', width: header.length });

    if (!res.rows.length) {
      rows.push(['', '(brak zawodników)']);
      fmt.push({ row: rows.length, kind: 'info' });
    }

    const firstBody = rows.length + 1;
    res.rows.forEach(function (a) {
      const r = [a.place, a.name, a.nr, a.bw === null ? '' : a.bw];
      a.comps.forEach(function (c) {
        r.push(c.best === null ? '' : c.best, c.score === null ? '' : c.score, c.place);
      });
      r.push(a.sum, a.tiebreak || '', a.points, a.notes.join('; '));
      rows.push(r);
    });
    if (res.rows.length) {
      fmt.push({ row: firstBody, kind: 'body', count: res.rows.length, width: header.length, scoreCols: scoreCols });
    }

    res.unresolved.forEach(function (u) {
      rows.push(['', '⚠ Remis na miejscu ' + u.place + ': ' + u.names.join(', ') + ' - zaznacz zwycięzcę dogrywki na liście startowej.']);
      fmt.push({ row: rows.length, kind: 'warn' });
    });
  });

  // Zapis jednym setValues.
  const width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
  const grid = rows.map(function (r) { return r.concat(new Array(width - r.length).fill('')); });

  sh.clear();
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  if (sh.getMaxRows() < grid.length) sh.insertRowsAfter(sh.getMaxRows(), grid.length - sh.getMaxRows());
  sh.getRange(1, 1, grid.length, width).setValues(grid).setVerticalAlignment('middle');

  fmt.forEach(function (f) {
    if (f.kind === 'title') sh.getRange(f.row, 1).setFontSize(16).setFontWeight('bold');
    if (f.kind === 'info') sh.getRange(f.row, 1, 1, 2).setFontColor('#666666').setFontStyle('italic');
    if (f.kind === 'section') sh.getRange(f.row, 1).setFontSize(13).setFontWeight('bold');
    if (f.kind === 'warn') sh.getRange(f.row, 2).setFontColor('#c00000').setFontWeight('bold');
    if (f.kind === 'header') {
      sh.getRange(f.row, 1, 1, f.width).setFontWeight('bold').setBackground('#d9e2f3').setWrap(true).setHorizontalAlignment('center');
    }
    if (f.kind === 'body') {
      const body = sh.getRange(f.row, 1, f.count, f.width);
      body.setHorizontalAlignment('center');
      sh.getRange(f.row, 2, f.count, 1).setHorizontalAlignment('left');
      sh.getRange(f.row, f.width, f.count, 1).setHorizontalAlignment('left').setFontColor('#666666');
      sh.getRange(f.row, 1, f.count, 1).setFontWeight('bold');
      sh.getRange(f.row, 4, f.count, 1).setNumberFormat('0.0');
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
      tieWon: r[meta.tieCol - 1] === true,
      attempts: meta.comps.map(function (c) {
        return r.slice(c.col - 1, c.col - 1 + c.attempts).map(toNum_).filter(function (v) { return v !== null; });
      }),
    });
  });
  return out;
}

/**
 * Liczy wyniki kategorii. Czysta funkcja (bez API arkusza) - łatwa do przetestowania.
 * athletes: [{ nr, name, bw, tieWon, attempts: [[kg, ...] per konkurencja] }]
 */
function computeRanking_(athletes, nComps, coef, decimals) {
  const rows = athletes.map(function (a) {
    const notes = [];
    if (a.bw === null) notes.push('brak wagi');
    const comps = [];
    for (let i = 0; i < nComps; i++) {
      const att = a.attempts[i] || [];
      const best = att.length ? Math.max.apply(null, att) : null;
      const score = best !== null && a.bw ? round_(best / Math.pow(a.bw, coef), decimals) : null;
      comps.push({ best: best, score: score, place: null });
    }
    return { nr: a.nr, name: a.name, bw: a.bw, tieWon: a.tieWon, comps: comps, notes: notes };
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

  // Dogrywka: zaznaczona wygrana zawsze daje -0,5 pkt.
  rows.forEach(function (r) {
    r.tiebreak = r.tieWon ? TIEBREAK_BONUS : 0;
    r.points = r.sum + r.tiebreak;
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
    sh.getRange(CFG.lockConfigCell).setValue(lockLabel_(false));
    props.setProperty(PROP_LOCK_CONFIG, '0');
  } else {
    sh.getRange(CFG.lockConfigCell).setValue(lockLabel_(true));
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
    if (locked) removeProtections_(x.sheet, SpreadsheetApp.ProtectionType.RANGE, 'zbm:zawodnicy');
    else protectAthletes_(x.sheet);
  });
  props.setProperty(PROP_LOCK_ATHLETES, locked ? '0' : '1');

  const cfgSheet = ss.getSheetByName(SHEET_CONFIG);
  if (cfgSheet) {
    try { cfgSheet.getRange(CFG.lockAthletesCell).setValue(lockLabel_(!locked)); } catch (e) { /* konfiguracja zablokowana dla tego użytkownika */ }
  }
  buildMenus_();
  ss.toast(locked ? 'Dane zawodników odblokowane.' : 'Dane zawodników zablokowane.', '🏋️ Zawody');
}

function protectAthletes_(sheet) {
  removeProtections_(sheet, SpreadsheetApp.ProtectionType.RANGE, 'zbm:zawodnicy');
  const p = sheet.getRange(LIST.firstRow, 2, LIST.maxRows, LIST.fixedCols - 1).protect().setDescription('zbm:zawodnicy');
  hardProtect_(p);
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
