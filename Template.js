/**
 * Szablon arkusza zawody-best-move (zawody hardstyle kettlebell).
 *
 * setupTemplate() jest idempotentna:
 *  - „Konfiguracja” (dane wpisywane przez Kacpra) - zakładana tylko, jeśli nie istnieje,
 *  - „Wyniki” (zakładka wyliczana) - przeliczana od nowa,
 *  - listy startowe tworzy się osobno: ⚙️ Setup → Generuj listy startowe.
 */

const SHEET_CONFIG = 'Konfiguracja';
const SHEET_RESULTS = 'Wyniki';
const LIST_PREFIX = 'Lista - ';

/** Dostępne konkurencje. Kolejność = kolejność kolumn w konfiguracji, listach i wynikach. */
const COMPETITIONS = [
  { key: 'PRESS', name: 'One Arm Press', short: 'Press' },
  { key: 'TGU', name: 'Turkish Get Up', short: 'TGU' },
];

/** Układ zakładki „Konfiguracja”. */
const CFG = {
  titleCell: 'B3',
  decimalsCell: 'B4',
  lockConfigCell: 'B5',
  lockAthletesCell: 'B6',
  // Tabela nazw konkurencji (D3:F5): etykieta, nazwa, skrót - wiersz na konkurencję w kolejności COMPETITIONS.
  compHeaderRow: 3,
  compFirstRow: 4,
  compLabelCol: 4,
  compNameCol: 5,
  compShortCol: 6,
  // Tabela kategorii: Kategoria, Współczynnik, Junior, potem dla każdej konkurencji: czy jest, ilość prób.
  headerRow: 8,
  firstRow: 9,
  maxRows: 8, // maksymalnie 8 kategorii
  // Tabela zawodnika referencyjnego: Kategoria (wybór z listy), dla każdej konkurencji waga ref., ciężar ref.
  // Powiązanie po NAZWIE kategorii - sortowanie którejkolwiek tabeli niczego nie psuje.
  refTitleRow: 19,
  refHeaderRow: 20,
  refFirstRow: 21,
  // Tabela walidacji: Reguła, Włączona, Wartość.
  valTitleRow: 31,
  valHeaderRow: 32,
  valFirstRow: 33,
};

/** Kolumna (1-based) w tabeli kategorii: konkurencja i (0-based), 0 = czy jest, 1 = ilość prób. */
function cfgCompCol_(i, offset) {
  return 4 + 2 * i + offset;
}

/** Kolumna (1-based) w tabeli referencji: konkurencja i (0-based), 0 = waga ref., 1 = ciężar ref. */
function cfgRefCol_(i, offset) {
  return 2 + 2 * i + offset;
}

/** Szerokość tabeli kategorii i tabeli referencji. */
function cfgCatCols_() { return 3 + 2 * COMPETITIONS.length; }
function cfgRefCols_() { return 1 + 2 * COMPETITIONS.length; }

/** Adres A1 komórki z nazwą (field = 'name') lub skrótem ('short') konkurencji i; withSheet = z nazwą zakładki. */
function compCellA1_(i, field, withSheet) {
  const col = field === 'short' ? CFG.compShortCol : CFG.compNameCol;
  const a1 = String.fromCharCode(64 + col) + (CFG.compFirstRow + i);
  return withSheet ? "'" + SHEET_CONFIG + "'!" + a1 : a1;
}

/** Konkurencje z nazwami z konfiguracji (puste pole = nazwa domyślna). */
function readCompetitions_(sh) {
  const vals = sh
    ? sh.getRange(CFG.compFirstRow, CFG.compNameCol, COMPETITIONS.length, 2).getValues()
    : COMPETITIONS.map(function () { return ['', '']; });
  return COMPETITIONS.map(function (c, i) {
    const name = String(vals[i][0]).trim();
    const short = String(vals[i][1]).trim();
    return { key: c.key, name: name || c.name, short: short || name || c.short };
  });
}

/** Kategoria, współczynnik, junior, potem dla każdej konkurencji: [czy jest, ilość prób]. */
const DEFAULT_CATEGORIES = [
  ['Kobiety', 0.60, false, true, 3, true, 3],
  ['Mężczyźni', 0.67, false, true, 3, true, 3],
  ['Juniorzy - chłopcy', 0.67, true, false, 3, true, 3],
  ['Juniorzy - dziewczęta', 0.60, true, false, 3, true, 3],
];

/** Zawodnik referencyjny: kategoria, Press [waga, ciężar], TGU [waga, ciężar]. */
const DEFAULT_REFS = [
  ['Kobiety', 65, 36, 65, 48],
  ['Mężczyźni', 85, 60, 85, 80],
  ['Juniorzy - chłopcy', 85, 60, 85, 80],
  ['Juniorzy - dziewczęta', 65, 36, 65, 48],
];

/** Reguły walidacji (kolejność = wiersze tabeli). value = domyślna wartość (tylko reguły z liczbą). */
const VALIDATION_RULES = [
  { key: 'WEIGHT', label: 'Waga zawodnika uzupełniona' },
  { key: 'ATTEMPTS', label: 'Wszystkie próby uzupełnione' },
  { key: 'ORDER', label: 'Każda próba cięższa lub równa poprzedniej' },
  { key: 'DUPLICATE', label: 'Brak powtórzonych zawodników w kategorii (imię + nazwisko)' },
  { key: 'TIES', label: 'Remisy na miejscach 1-3 rozstrzygnięte dogrywką' },
  { key: 'CONFIG', label: 'Listy i referencje zgodne z konfiguracją' },
  { key: 'LIMIT_TOTAL', label: 'Limit zawodników łącznie', value: 40 },
  { key: 'LIMIT_ADULT', label: 'Limit zawodników dorosłych', value: 30 },
  { key: 'LIMIT_JUNIOR', label: 'Limit juniorów (kategorie z zaznaczonym „Junior”)', value: 10 },
];

const DEFAULT_DECIMALS = 3;

function onOpen() {
  buildMenus_();
}

function buildMenus_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Setup')
    .addItem('Zbuduj szablon', 'setupTemplate')
    .addItem('Generuj listy startowe', 'generateStartLists')
    .addToUi();
  buildZawodyMenu_(ui);
  ui.createMenu('🖨️ Druk')
    .addItem('Drukuj regulamin (PDF)', 'printRules')
    .addItem('Drukuj listę zawodników (PDF)', 'printAthletes')
    .addItem('Drukuj wyniki (PDF)', 'printResults')
    .addToUi();
}

function setupTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_CONFIG)) buildConfigSheet_(ss);
  if (!ss.getSheetByName(SHEET_RESULTS)) createResultsSheet_(ss);
  if (!ss.getSheetByName(SHEET_RULES)) buildRulesSheet_(ss);
  rebuildResults();

  // Domyślna pusta zakładka nowego arkusza - usuwana tylko, gdy naprawdę pusta.
  ['Arkusz1', 'Sheet1'].forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() === 0 && sh.getLastColumn() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
    }
  });

  ss.setActiveSheet(ss.getSheetByName(SHEET_CONFIG));
  ss.toast('Szablon zbudowany. Dalej: ⚙️ Setup → Generuj listy startowe.', '⚙️ Setup', 8);
}

function buildConfigSheet_(ss) {
  const sh = ss.insertSheet(SHEET_CONFIG, 0);
  const nCols = cfgCatCols_();
  const nRefCols = cfgRefCols_();
  const GRID = '#c0c7d4';
  const note = function (row, text) {
    sh.getRange(row, 1).setValue(text).setFontColor('#666666').setFontStyle('italic');
  };
  const headerStyle = function (range) {
    return range.setFontWeight('bold').setBackground('#d9e2f3').setWrap(true).setVerticalAlignment('middle');
  };
  const grid = function (range) {
    range.setBorder(true, true, true, true, true, true, GRID, SpreadsheetApp.BorderStyle.SOLID);
  };

  sh.getRange('A1').setValue('Konfiguracja zawodów').setFontSize(14).setFontWeight('bold');

  sh.getRange('A3:A6').setValues([
    ['Nazwa zawodów'],
    ['Miejsca po przecinku (wynik)'],
    ['Blokada konfiguracji'],
    ['Blokada danych zawodników'],
  ]).setFontWeight('bold');
  sh.getRange(CFG.decimalsCell)
    .setValue(DEFAULT_DECIMALS)
    .setNumberFormat('0')
    .setHorizontalAlignment('left')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 6).setAllowInvalid(false)
      .setHelpText('Liczba całkowita 0-6').build());
  sh.getRange('B5:B6').setValues([[lockLabel_(false)], [lockLabel_(false)]]).setFontColor('#666666');

  // ── Nazwy konkurencji ──
  headerStyle(sh.getRange(CFG.compHeaderRow, CFG.compLabelCol, 1, 3).setValues([['Konkurencja', 'Nazwa konkurencji', 'Skrót']]));
  sh.getRange(CFG.compFirstRow, CFG.compLabelCol, COMPETITIONS.length, 3)
    .setValues(COMPETITIONS.map(function (c, i) { return ['Konkurencja ' + (i + 1), c.name, c.short]; }));
  sh.getRange(CFG.compFirstRow, CFG.compLabelCol, COMPETITIONS.length, 1).setFontColor('#666666');
  grid(sh.getRange(CFG.compHeaderRow, CFG.compLabelCol, COMPETITIONS.length + 1, 3));

  // ── Kategorie ── (nagłówki konkurencji jako formuły - zmiana nazwy od razu widoczna)
  const headers = ['Kategoria', 'Współczynnik (0-1)', 'Junior'];
  COMPETITIONS.forEach(function (c) { headers.push(c.name, 'Próby - ' + c.short); });
  headerStyle(sh.getRange(CFG.headerRow, 1, 1, nCols).setValues([headers]));
  COMPETITIONS.forEach(function (c, i) {
    sh.getRange(CFG.headerRow, cfgCompCol_(i, 0)).setFormula('=' + compCellA1_(i, 'name'));
    sh.getRange(CFG.headerRow, cfgCompCol_(i, 1)).setFormula('="Próby - "&' + compCellA1_(i, 'short'));
  });

  const body = function (col) { return sh.getRange(CFG.firstRow, col, CFG.maxRows, 1); };
  body(2).setNumberFormat('0.00').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 1).setAllowInvalid(false)
    .setHelpText('Współczynnik od 0 do 1, np. 0,67').build());
  body(3).insertCheckboxes().setHorizontalAlignment('center');
  COMPETITIONS.forEach(function (c, i) {
    body(cfgCompCol_(i, 0)).insertCheckboxes().setHorizontalAlignment('center');
    body(cfgCompCol_(i, 1)).setNumberFormat('0').setHorizontalAlignment('center')
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireNumberBetween(1, 10).setAllowInvalid(false)
        .setHelpText('Ilość prób: 1-10').build());
  });
  sh.getRange(CFG.firstRow, 1, DEFAULT_CATEGORIES.length, nCols).setValues(DEFAULT_CATEGORIES);
  grid(sh.getRange(CFG.headerRow, 1, CFG.maxRows + 1, nCols));
  note(CFG.firstRow + CFG.maxRows,
    'Maksymalnie ' + CFG.maxRows + ' kategorii. Pusty wiersz = brak kategorii. „Junior” = liczy się do limitu juniorów. Zmiana konkurencji lub ilości prób działa dla nowo generowanych list.');

  // ── Zawodnik referencyjny ──
  sh.getRange(CFG.refTitleRow, 1).setValue('Zawodnik referencyjny').setFontSize(12).setFontWeight('bold');
  const refHeaders = ['Kategoria'];
  COMPETITIONS.forEach(function (c) { refHeaders.push(c.short + ' - waga zawodnika ref. [kg]', c.short + ' - ciężar ref. [kg]'); });
  headerStyle(sh.getRange(CFG.refHeaderRow, 1, 1, nRefCols).setValues([refHeaders]));
  COMPETITIONS.forEach(function (c, i) {
    sh.getRange(CFG.refHeaderRow, cfgRefCol_(i, 0)).setFormula('=' + compCellA1_(i, 'short') + '&" - waga zawodnika ref. [kg]"');
    sh.getRange(CFG.refHeaderRow, cfgRefCol_(i, 1)).setFormula('=' + compCellA1_(i, 'short') + '&" - ciężar ref. [kg]"');
  });
  sh.getRange(CFG.refFirstRow, 1, CFG.maxRows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getRange(CFG.firstRow, 1, CFG.maxRows, 1), true).setAllowInvalid(false)
    .setHelpText('Wybierz kategorię z tabeli kategorii.').build());
  sh.getRange(CFG.refFirstRow, 2, CFG.maxRows, nRefCols - 1)
    .setNumberFormat('0.0').setHorizontalAlignment('center')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0).setAllowInvalid(false)
      .setHelpText('Wartość w kg (dokładność 0,1)').build());
  sh.getRange(CFG.refFirstRow, 1, DEFAULT_REFS.length, nRefCols).setValues(DEFAULT_REFS);
  grid(sh.getRange(CFG.refHeaderRow, 1, CFG.maxRows + 1, nRefCols));
  note(CFG.refFirstRow + CFG.maxRows,
    'Wiersz łączy się z kategorią po nazwie. wynik = ciężar / waga^wsp * waga_ref^wsp / ciężar_ref - zawodnik referencyjny ma wynik 1. Brak wiersza = wynik bez skalowania.');

  // ── Walidacja ──
  sh.getRange(CFG.valTitleRow, 1).setValue('Walidacja zawodów').setFontSize(12).setFontWeight('bold');
  headerStyle(sh.getRange(CFG.valHeaderRow, 1, 1, 3).setValues([['Reguła', 'Włączona', 'Wartość']]));
  const nVal = VALIDATION_RULES.length;
  sh.getRange(CFG.valFirstRow, 2, nVal, 1).insertCheckboxes().setHorizontalAlignment('center');
  sh.getRange(CFG.valFirstRow, 1, nVal, 3).setValues(VALIDATION_RULES.map(function (r) {
    return [r.label, true, r.value === undefined ? '' : r.value];
  }));
  VALIDATION_RULES.forEach(function (r, i) {
    const cell = sh.getRange(CFG.valFirstRow + i, 3).setHorizontalAlignment('center');
    if (r.value === undefined) {
      cell.setBackground('#f3f3f3');
    } else {
      cell.setNumberFormat('0').setDataValidation(SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false)
        .setHelpText('Liczba zawodników').build());
    }
  });
  sh.getRange(CFG.valFirstRow, 1, nVal, 1).setWrap(true);
  grid(sh.getRange(CFG.valHeaderRow, 1, nVal + 1, 3));
  note(CFG.valFirstRow + nVal, 'Sprawdzane przez 🏋️ Zawody → ✅ Walidacja zawodów oraz przed drukiem wyników.');

  sh.setColumnWidth(1, 300);
  for (let c = 2; c <= Math.max(nCols, nRefCols, CFG.compShortCol); c++) sh.setColumnWidth(c, 130);
  return sh;
}

/** Odczytuje konfigurację zawodów. Referencje łączone z kategoriami po nazwie. */
function readConfig_(ss) {
  const sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) throw new Error('Brak zakładki „' + SHEET_CONFIG + '” - uruchom ⚙️ Setup → Zbuduj szablon.');

  const rows = sh.getRange(CFG.firstRow, 1, CFG.maxRows, cfgCatCols_()).getValues();
  const refRows = sh.getRange(CFG.refFirstRow, 1, CFG.maxRows, cfgRefCols_()).getValues();
  const valRows = sh.getRange(CFG.valFirstRow, 1, VALIDATION_RULES.length, 3).getValues();

  const decRaw = sh.getRange(CFG.decimalsCell).getValue();
  const decimals = decRaw === '' || isNaN(Number(decRaw))
    ? DEFAULT_DECIMALS
    : Math.max(0, Math.min(6, Math.round(Number(decRaw))));

  const parsed = parseConfigRows_(rows, refRows, readCompetitions_(sh));
  return {
    title: String(sh.getRange(CFG.titleCell).getValue()).trim(),
    decimals: decimals,
    competitions: parsed.competitions,
    categories: parsed.categories,
    configIssues: parsed.issues,
    validation: parseValidationRows_(valRows),
  };
}

/** Czysta funkcja: wiersze kategorii + referencji -> kategorie i problemy konfiguracji. */
function parseConfigRows_(rows, refRows, competitions) {
  const norm = function (s) { return String(s).trim().toLowerCase(); };
  const num = function (v) { return v === '' || v === null ? null : Number(v); };
  const issues = [];

  const refs = {};
  refRows.forEach(function (r) {
    const key = norm(r[0]);
    if (!key) return;
    if (refs[key]) issues.push('Tabela referencji: kategoria „' + String(r[0]).trim() + '” wpisana więcej niż raz (liczy się pierwszy wiersz)');
    else refs[key] = r;
  });

  const categories = [];
  const seen = {};
  rows.forEach(function (r) {
    const name = String(r[0]).trim();
    if (!name) return;
    if (seen[norm(name)]) { issues.push('Kategoria „' + name + '” wpisana więcej niż raz'); return; }
    seen[norm(name)] = true;
    const ref = refs[norm(name)];
    if (!ref) issues.push('Kategoria „' + name + '”: brak wiersza w tabeli zawodnika referencyjnego (wynik bez skalowania)');
    const comps = [];
    competitions.forEach(function (c, i) {
      if (r[cfgCompCol_(i, 0) - 1] === true) {
        const attempts = Math.round(Number(r[cfgCompCol_(i, 1) - 1]));
        comps.push({
          key: c.key, name: c.name, short: c.short,
          attempts: attempts >= 1 ? attempts : 3,
          refBodyweight: ref ? num(ref[cfgRefCol_(i, 0) - 1]) : null,
          refWeight: ref ? num(ref[cfgRefCol_(i, 1) - 1]) : null,
        });
      }
    });
    categories.push({ name: name, coef: r[1] === '' ? NaN : Number(r[1]), junior: r[2] === true, comps: comps });
  });

  Object.keys(refs).forEach(function (key) {
    if (!seen[key]) issues.push('Tabela referencji: „' + String(refs[key][0]).trim() + '” nie jest kategorią z tabeli kategorii');
  });

  return { competitions: competitions, categories: categories, issues: issues };
}

/** Czysta funkcja: wiersze tabeli walidacji -> { KLUCZ: { on, value } }. Brak wiersza = reguła włączona. */
function parseValidationRows_(valRows) {
  const out = {};
  VALIDATION_RULES.forEach(function (rule, i) {
    const r = valRows[i] || [];
    const value = r[2] === '' || r[2] === undefined || isNaN(Number(r[2])) ? (rule.value === undefined ? null : rule.value) : Number(r[2]);
    out[rule.key] = { on: r[1] === undefined || r[1] === '' ? true : r[1] === true, value: value };
  });
  return out;
}

function createResultsSheet_(ss) {
  const sh = ss.insertSheet(SHEET_RESULTS, ss.getSheets().length);
  sh.protect().setDescription('zbm:wyniki').setWarningOnly(true);
  sh.setColumnWidth(1, 70);
  sh.setColumnWidth(2, 220);
  return sh;
}

function lockLabel_(locked) {
  return locked ? '🔒 zablokowana' : '🔓 odblokowana';
}

/** Zwraca zakładkę o danej nazwie, zakładając ją, jeśli nie istnieje (bez czyszczenia danych). */
function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
