/**
 * Ochrona pól, których nie edytuje się ręcznie (nagłówki, etykiety, opisy, statusy).
 * Ochrona Google: działa na edytorów bez uprawnień właściciela (osoby wpisujące dane).
 * Idempotentna - przy każdym wywołaniu usuwa stare ochrony „zbm:stale” danej zakładki i zakłada nowe.
 */

const STATIC_DESC = 'zbm:stale';

/** Litera kolumny (1 = A, 27 = AA). */
function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Czysta funkcja: zakresy A1 stałych pól Konfiguracji. */
function configStaticRanges_() {
  const L = colLetter_;
  const n = VALIDATION_RULES.length;
  const wide = L(Math.max(cfgCatCols_(), cfgRefCols_(), CFG.compShortCol));
  const catEnd = L(cfgCatCols_());
  const refEnd = L(cfgRefCols_());
  const compStart = L(CFG.compLabelCol), compEnd = L(CFG.compShortCol);
  const noteCat = CFG.firstRow + CFG.maxRows;
  const noteRef = CFG.refFirstRow + CFG.maxRows;
  const valNote = CFG.valFirstRow + n;

  const ranges = [
    'A1:' + wide + '2',                                                    // tytuł
    'A3:A7',                                                               // etykiety ustawień
    'B5:B6',                                                               // status blokad (wpisuje skrypt)
    compStart + CFG.compHeaderRow + ':' + compEnd + CFG.compHeaderRow,     // nagłówek nazw konkurencji
    compStart + CFG.compFirstRow + ':' + compStart + (CFG.compFirstRow + COMPETITIONS.length - 1), // „Konkurencja 1/2”
    'A' + CFG.headerRow + ':' + catEnd + CFG.headerRow,                    // nagłówek kategorii
    'A' + noteCat + ':' + wide + (CFG.refTitleRow - 1),                    // notatka pod kategoriami
    'A' + CFG.refTitleRow + ':' + refEnd + CFG.refHeaderRow,               // tytuł + nagłówek referencji
    'A' + noteRef + ':' + wide + (CFG.valTitleRow - 1),                    // notatka pod referencjami
    'A' + CFG.valTitleRow + ':C' + CFG.valHeaderRow,                       // tytuł + nagłówek walidacji
    'A' + CFG.valFirstRow + ':A' + (CFG.valFirstRow + n - 1),              // opisy reguł
    'A' + valNote + ':' + wide + valNote,                                  // notatka pod walidacją
  ];
  VALIDATION_RULES.forEach(function (r, i) {
    if (r.value === undefined) ranges.push('C' + (CFG.valFirstRow + i)); // reguła bez wartości
  });
  return ranges;
}

/** Zakłada ochronę na podanych zakresach (usuwa wcześniejsze ochrony stałych pól tej zakładki). */
function protectStatic_(sheet, a1List) {
  removeProtections_(sheet, SpreadsheetApp.ProtectionType.RANGE, STATIC_DESC);
  a1List.forEach(function (a1) {
    hardProtect_(sheet.getRange(a1).protect().setDescription(STATIC_DESC));
  });
}

/** Lista startowa: wiersze 1-3 (tytuł, nazwy konkurencji, nagłówki kolumn). Nr chroniony osobno. */
function protectListStatic_(sheet) {
  protectStatic_(sheet, ['A1:' + colLetter_(sheet.getMaxColumns()) + LIST.headerRow]);
}

/** Regulamin: wiersz nagłówka. */
function protectRulesStatic_(sheet) {
  protectStatic_(sheet, ['A' + RULES.headerRow + ':B' + RULES.headerRow]);
}

/** Wszystkie istniejące zakładki (wywoływane z „Zbuduj szablon”). */
function applyStaticProtections_(ss) {
  const cfg = ss.getSheetByName(SHEET_CONFIG);
  if (cfg) protectStatic_(cfg, configStaticRanges_());
  getAllListSheets_(ss).forEach(function (x) { protectListStatic_(x.sheet); });
  const rules = ss.getSheetByName(SHEET_RULES);
  if (rules) protectRulesStatic_(rules);
}
