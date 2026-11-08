/**
 * Szablon arkusza zawody-best-move (zawody hardstyle kettlebell).
 *
 * setupTemplate() jest idempotentna:
 *  - „Konfiguracja” (dane wpisywane przez Kacpra) - zakładana tylko, jeśli nie istnieje,
 *  - „Wyniki” (zakładka wyliczana) - przeliczana od nowa,
 *  - listy startowe tworzy się osobno: 🏋️ Zawody → Generuj listy startowe.
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
  // Tabela kategorii: Kategoria, Współczynnik, potem dla każdej konkurencji: czy jest, ilość prób.
  headerRow: 8,
  firstRow: 9,
  maxRows: 8, // maksymalnie 8 kategorii
  // Tabela zawodnika referencyjnego (wiersze = te same kategorie): dla każdej konkurencji waga ref., ciężar ref.
  refTitleRow: 19,
  refHeaderRow: 20,
  refFirstRow: 21,
};

/** Kolumna (1-based) w tabeli kategorii: konkurencja i (0-based), 0 = czy jest, 1 = ilość prób. */
function cfgCompCol_(i, offset) {
  return 3 + 2 * i + offset;
}

/** Kolumna (1-based) w tabeli referencji: konkurencja i (0-based), 0 = waga ref., 1 = ciężar ref. */
function cfgRefCol_(i, offset) {
  return 2 + 2 * i + offset;
}

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

/** Kategoria, współczynnik, potem dla każdej konkurencji: [czy jest, ilość prób]. */
const DEFAULT_CATEGORIES = [
  ['Kobiety', 0.60, true, 3, true, 3],
  ['Mężczyźni', 0.67, true, 3, true, 3],
  ['Juniorzy - chłopcy', 0.67, false, 3, true, 3],
  ['Juniorzy - dziewczęta', 0.60, false, 3, true, 3],
];

/** Zawodnik referencyjny (te same wiersze co kategorie): Press [waga, ciężar], TGU [waga, ciężar]. */
const DEFAULT_REFS = [
  [65, 36, 65, 48],
  [85, 60, 85, 80],
  [85, 60, 85, 80],
  [65, 36, 65, 48],
];

const DEFAULT_DECIMALS = 3;

function onOpen() {
  buildMenus_();
}

function buildMenus_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Setup')
    .addItem('Zbuduj szablon', 'setupTemplate')
    .addToUi();
  buildZawodyMenu_(ui);
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
  ss.toast('Szablon zbudowany. Dalej: 🏋️ Zawody → Generuj listy startowe.', '⚙️ Setup', 8);
}

function buildConfigSheet_(ss) {
  const sh = ss.insertSheet(SHEET_CONFIG, 0);
  const nCols = 2 + 2 * COMPETITIONS.length;
  const nRefCols = 1 + 2 * COMPETITIONS.length;
  const note = function (row, text) {
    sh.getRange(row, 1).setValue(text).setFontColor('#666666').setFontStyle('italic');
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

  // Tabela nazw konkurencji
  sh.getRange(CFG.compHeaderRow, CFG.compLabelCol, 1, 3)
    .setValues([['Konkurencja', 'Nazwa konkurencji', 'Skrót']])
    .setFontWeight('bold').setBackground('#d9e2f3');
  sh.getRange(CFG.compFirstRow, CFG.compLabelCol, COMPETITIONS.length, 3)
    .setValues(COMPETITIONS.map(function (c, i) { return ['Konkurencja ' + (i + 1), c.name, c.short]; }));
  sh.getRange(CFG.compFirstRow, CFG.compLabelCol, COMPETITIONS.length, 1).setFontColor('#666666');
  sh.getRange(CFG.compHeaderRow, CFG.compLabelCol, COMPETITIONS.length + 1, 3)
    .setBorder(true, true, true, true, true, true, '#c0c7d4', SpreadsheetApp.BorderStyle.SOLID);

  // Tabela kategorii - nagłówki konkurencji jako formuły (zmiana nazwy od razu widoczna).
  const headers = ['Kategoria', 'Współczynnik (0-1)'];
  COMPETITIONS.forEach(function (c) { headers.push(c.name, 'Próby - ' + c.short); });
  sh.getRange(CFG.headerRow, 1, 1, nCols).setValues([headers])
    .setFontWeight('bold').setBackground('#d9e2f3').setWrap(true).setVerticalAlignment('middle');
  COMPETITIONS.forEach(function (c, i) {
    sh.getRange(CFG.headerRow, cfgCompCol_(i, 0)).setFormula('=' + compCellA1_(i, 'name'));
    sh.getRange(CFG.headerRow, cfgCompCol_(i, 1)).setFormula('="Próby - "&' + compCellA1_(i, 'short'));
  });

  const body = function (col) { return sh.getRange(CFG.firstRow, col, CFG.maxRows, 1); };
  body(2).setNumberFormat('0.00').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 1).setAllowInvalid(false)
    .setHelpText('Współczynnik od 0 do 1, np. 0,67').build());
  COMPETITIONS.forEach(function (c, i) {
    body(cfgCompCol_(i, 0)).insertCheckboxes().setHorizontalAlignment('center');
    body(cfgCompCol_(i, 1)).setNumberFormat('0').setHorizontalAlignment('center')
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireNumberBetween(1, 10).setAllowInvalid(false)
        .setHelpText('Ilość prób: 1-10').build());
  });
  sh.getRange(CFG.firstRow, 1, DEFAULT_CATEGORIES.length, nCols).setValues(DEFAULT_CATEGORIES);
  sh.getRange(CFG.headerRow, 1, CFG.maxRows + 1, nCols)
    .setBorder(true, true, true, true, true, true, '#c0c7d4', SpreadsheetApp.BorderStyle.SOLID);
  note(CFG.firstRow + CFG.maxRows,
    'Maksymalnie ' + CFG.maxRows + ' kategorii. Pusty wiersz = brak kategorii. Zmiana konkurencji lub ilości prób działa dla nowo generowanych list - istniejącą listę trzeba usunąć i wygenerować od nowa.');

  // Tabela zawodnika referencyjnego
  sh.getRange(CFG.refTitleRow, 1).setValue('Zawodnik referencyjny').setFontSize(12).setFontWeight('bold');
  const refHeaders = ['Kategoria'];
  COMPETITIONS.forEach(function (c) { refHeaders.push(c.short + ' - waga zawodnika ref. [kg]', c.short + ' - ciężar ref. [kg]'); });
  sh.getRange(CFG.refHeaderRow, 1, 1, nRefCols).setValues([refHeaders])
    .setFontWeight('bold').setBackground('#d9e2f3').setWrap(true).setVerticalAlignment('middle');
  COMPETITIONS.forEach(function (c, i) {
    sh.getRange(CFG.refHeaderRow, cfgRefCol_(i, 0)).setFormula('=' + compCellA1_(i, 'short') + '&" - waga zawodnika ref. [kg]"');
    sh.getRange(CFG.refHeaderRow, cfgRefCol_(i, 1)).setFormula('=' + compCellA1_(i, 'short') + '&" - ciężar ref. [kg]"');
  });

  // Nazwy kategorii przepisywane z tabeli kategorii (nie wpisuje się ich drugi raz).
  const nameFormulas = [];
  for (let k = 0; k < CFG.maxRows; k++) {
    const src = 'A' + (CFG.firstRow + k);
    nameFormulas.push(['=IF(' + src + '="";"";' + src + ')']);
  }
  sh.getRange(CFG.refFirstRow, 1, CFG.maxRows, 1).setFormulas(nameFormulas).setFontColor('#666666').setBackground('#f3f3f3');

  sh.getRange(CFG.refFirstRow, 2, CFG.maxRows, nRefCols - 1)
    .setNumberFormat('0.0').setHorizontalAlignment('center')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0).setAllowInvalid(false)
      .setHelpText('Wartość w kg (dokładność 0,1)').build());
  sh.getRange(CFG.refFirstRow, 2, DEFAULT_REFS.length, nRefCols - 1).setValues(DEFAULT_REFS);
  sh.getRange(CFG.refHeaderRow, 1, CFG.maxRows + 1, nRefCols)
    .setBorder(true, true, true, true, true, true, '#c0c7d4', SpreadsheetApp.BorderStyle.SOLID);
  note(CFG.refFirstRow + CFG.maxRows,
    'wynik = ciężar / waga^wsp * waga_ref^wsp / ciężar_ref - zawodnik referencyjny ma wynik 1. Puste pola = wynik bez skalowania.');

  sh.setColumnWidth(1, 230);
  for (let c = 2; c <= Math.max(nCols, nRefCols); c++) sh.setColumnWidth(c, 130);
  return sh;
}

/** Odczytuje konfigurację zawodów (tabela kategorii + tabela referencji, wiersz w wiersz). */
function readConfig_(ss) {
  const sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) throw new Error('Brak zakładki „' + SHEET_CONFIG + '” - uruchom ⚙️ Setup → Zbuduj szablon.');

  const rows = sh.getRange(CFG.firstRow, 1, CFG.maxRows, 2 + 2 * COMPETITIONS.length).getValues();
  const refRows = sh.getRange(CFG.refFirstRow, 1, CFG.maxRows, 1 + 2 * COMPETITIONS.length).getValues();

  const decRaw = sh.getRange(CFG.decimalsCell).getValue();
  const decimals = decRaw === '' || isNaN(Number(decRaw))
    ? DEFAULT_DECIMALS
    : Math.max(0, Math.min(6, Math.round(Number(decRaw))));

  const num = function (v) { return v === '' || v === null ? null : Number(v); };
  const competitions = readCompetitions_(sh);
  const categories = [];
  rows.forEach(function (r, k) {
    const name = String(r[0]).trim();
    if (!name) return;
    const ref = refRows[k];
    const comps = [];
    competitions.forEach(function (c, i) {
      if (r[cfgCompCol_(i, 0) - 1] === true) {
        const attempts = Math.round(Number(r[cfgCompCol_(i, 1) - 1]));
        comps.push({
          key: c.key, name: c.name, short: c.short,
          attempts: attempts >= 1 ? attempts : 3,
          refBodyweight: num(ref[cfgRefCol_(i, 0) - 1]),
          refWeight: num(ref[cfgRefCol_(i, 1) - 1]),
        });
      }
    });
    categories.push({ name: name, coef: r[1] === '' ? NaN : Number(r[1]), comps: comps });
  });

  return {
    title: String(sh.getRange(CFG.titleCell).getValue()).trim(),
    decimals: decimals,
    competitions: competitions,
    categories: categories,
  };
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
