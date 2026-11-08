/**
 * Szablon arkusza zawody-best-move (zawody hardstyle kettlebell).
 *
 * setupTemplate() jest idempotentna:
 *  - „Konfiguracja” (dane wpisywane przez Kacpra) — zakładana tylko, jeśli nie istnieje,
 *  - „Wyniki” (zakładka wyliczana) — przeliczana od nowa,
 *  - listy startowe tworzy się osobno: 🏋️ Zawody → Generuj listy startowe.
 */

const SHEET_CONFIG = 'Konfiguracja';
const SHEET_RESULTS = 'Wyniki';
const LIST_PREFIX = 'Lista – ';

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
  lockWeightsCell: 'B6',
  headerRow: 8,
  firstRow: 9,
  maxRows: 20,
};

/** Kategoria, współczynnik, potem dla każdej konkurencji: [czy jest, ilość prób]. */
const DEFAULT_CATEGORIES = [
  ['Kobiety', 0.60, true, 3, true, 3],
  ['Mężczyźni', 0.67, true, 3, true, 3],
  ['Juniorzy – chłopcy', 0.67, false, 3, true, 3],
  ['Juniorzy – dziewczęta', 0.60, false, 3, true, 3],
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
  rebuildResults();

  // Domyślna pusta zakładka nowego arkusza — usuwana tylko, gdy naprawdę pusta.
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

  sh.getRange('A1').setValue('Konfiguracja zawodów').setFontSize(14).setFontWeight('bold');

  sh.getRange('A3:A6').setValues([
    ['Nazwa zawodów'],
    ['Miejsca po przecinku (wynik)'],
    ['Blokada konfiguracji'],
    ['Blokada wag'],
  ]).setFontWeight('bold');
  sh.getRange(CFG.decimalsCell)
    .setValue(DEFAULT_DECIMALS)
    .setNumberFormat('0')
    .setHorizontalAlignment('left')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 6).setAllowInvalid(false)
      .setHelpText('Liczba całkowita 0–6').build());
  sh.getRange('B5:B6').setValues([[lockLabel_(false)], [lockLabel_(false)]]).setFontColor('#666666');

  const headers = ['Kategoria', 'Współczynnik (0–1)'];
  COMPETITIONS.forEach(function (c) { headers.push(c.name, 'Próby – ' + c.short); });
  sh.getRange(CFG.headerRow, 1, 1, nCols).setValues([headers])
    .setFontWeight('bold').setBackground('#d9e2f3').setWrap(true).setVerticalAlignment('middle');

  const body = function (col) { return sh.getRange(CFG.firstRow, col, CFG.maxRows, 1); };
  body(2).setNumberFormat('0.00').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 1).setAllowInvalid(false)
    .setHelpText('Współczynnik od 0 do 1, np. 0,67').build());
  COMPETITIONS.forEach(function (c, i) {
    body(3 + 2 * i).insertCheckboxes().setHorizontalAlignment('center');
    body(4 + 2 * i).setNumberFormat('0').setHorizontalAlignment('center')
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireNumberBetween(1, 10).setAllowInvalid(false)
        .setHelpText('Ilość prób: 1–10').build());
  });

  sh.getRange(CFG.firstRow, 1, DEFAULT_CATEGORIES.length, nCols).setValues(DEFAULT_CATEGORIES);

  sh.getRange(CFG.firstRow + CFG.maxRows + 1, 1)
    .setValue('Pusty wiersz = brak kategorii. Zmiana konkurencji lub ilości prób działa dla nowo generowanych list — istniejącą listę trzeba usunąć i wygenerować od nowa.')
    .setFontColor('#666666').setFontStyle('italic');

  sh.setColumnWidth(1, 230);
  for (let c = 2; c <= nCols; c++) sh.setColumnWidth(c, 120);
  sh.setFrozenRows(CFG.headerRow);
  return sh;
}

/** Odczytuje konfigurację zawodów. */
function readConfig_(ss) {
  const sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) throw new Error('Brak zakładki „' + SHEET_CONFIG + '” — uruchom ⚙️ Setup → Zbuduj szablon.');

  const nCols = 2 + 2 * COMPETITIONS.length;
  const rows = sh.getRange(CFG.firstRow, 1, CFG.maxRows, nCols).getValues();

  const decRaw = sh.getRange(CFG.decimalsCell).getValue();
  const decimals = decRaw === '' || isNaN(Number(decRaw))
    ? DEFAULT_DECIMALS
    : Math.max(0, Math.min(6, Math.round(Number(decRaw))));

  const categories = [];
  rows.forEach(function (r) {
    const name = String(r[0]).trim();
    if (!name) return;
    const comps = [];
    COMPETITIONS.forEach(function (c, i) {
      if (r[2 + 2 * i] === true) {
        const attempts = Math.round(Number(r[3 + 2 * i]));
        comps.push({ key: c.key, name: c.name, short: c.short, attempts: attempts >= 1 ? attempts : 3 });
      }
    });
    categories.push({ name: name, coef: r[1] === '' ? NaN : Number(r[1]), comps: comps });
  });

  return {
    title: String(sh.getRange(CFG.titleCell).getValue()).trim(),
    decimals: decimals,
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
