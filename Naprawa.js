/**
 * Naprawa list startowych: checkboxy „Zal.”, reguły (ciężar, waga, dogrywka), formaty i ramki.
 * Checkbox to reguła poprawności danych - Delete na kratce albo wklejenie skopiowanych komórek ją usuwa.
 *  - onEdit: po każdej edycji listy sprawdza edytowane wiersze i przywraca brakujące reguły (wartości zostają),
 *  - ⚙️ Setup → Napraw listy startowe (i „Zbuduj szablon”): pełna naprawa wszystkich list bez ruszania danych.
 */

/** Czysta funkcja: które kolumny listy mają jaką regułę. */
function listRuleSpec_(meta) {
  const spec = [{ col: LIST.weightCol, kind: 'weight' }];
  if (meta.layout === 2) {
    meta.comps.forEach(function (c) {
      for (let i = 0; i < c.attempts; i++) {
        spec.push({ col: c.col + 2 * i, kind: 'kg' }, { col: c.col + 2 * i + 1, kind: 'check' });
      }
    });
  }
  spec.push({ col: meta.tieCol, kind: 'tie' });
  return spec;
}

function listRule_(kind) {
  const b = SpreadsheetApp.newDataValidation().setAllowInvalid(false);
  if (kind === 'weight') {
    return b.requireNumberGreaterThan(0).setHelpText('Waga zawodnika w kg (dokładność 0,1).').build();
  }
  if (kind === 'kg') {
    return b.requireNumberGreaterThanOrEqualTo(0)
      .setHelpText('Ciężar w kg (dokładność 0,1). Próba zaliczona = zaznacz „Zal.” obok. Spalona o nieznanym ciężarze = 0.').build();
  }
  return b.requireValueInList(Object.keys(TIEBREAK_BONUS), true)
    .setHelpText('Miejsce w dogrywce: 1 = -0,5 pkt, 2 = -0,33 pkt, 3 = -0,25 pkt. Puste = bez dogrywki.').build();
}

/**
 * Reguły i formaty kolumn listy w wierszach r1..r1+n-1 (wartości zostają).
 * onlyCols - opcjonalnie tylko te kolumny; każda kolumna osobno (chroniona kolumna nie blokuje pozostałych).
 */
function applyListRules_(sh, meta, r1, n, onlyCols) {
  listRuleSpec_(meta).forEach(function (s) {
    if (onlyCols && onlyCols.indexOf(s.col) < 0) return;
    try {
      const range = sh.getRange(r1, s.col, n, 1).setHorizontalAlignment('center');
      if (s.kind === 'check') range.insertCheckboxes();
      else range.setDataValidation(listRule_(s.kind));
      if (s.kind === 'weight' || s.kind === 'kg') range.setNumberFormat('0.0');
    } catch (err) {
      console.error('Naprawa kolumny ' + s.col + ': ' + err); // np. kolumna chroniona dla tego użytkownika
    }
  });
  if (meta.notesCol) sh.getRange(r1, meta.notesCol, n, 1).setWrap(true);
}

/** Ramki listy: siatka, para [ciężar | Zal.], grube linie przy konkurencjach i dogrywce. */
function applyListBorders_(sh, meta) {
  const lastRow = LIST.firstRow + LIST.maxRows - 1;
  const nCols = listWidth_(meta);
  const SB = SpreadsheetApp.BorderStyle;
  sh.getRange(LIST.groupRow, 1, lastRow - LIST.groupRow + 1, nCols)
    .setBorder(true, true, true, true, true, true, '#c0c7d4', SB.SOLID);
  const H = LIST.headerRow, nBody = lastRow - LIST.headerRow + 1;
  if (meta.layout === 2) {
    meta.comps.forEach(function (c) {
      for (let i = 0; i < c.attempts; i++) {
        sh.getRange(H, c.col + 2 * i, nBody, 2).setBorder(null, true, null, true, null, null, '#8896ab', SB.SOLID_MEDIUM);
      }
      sh.getRange(LIST.groupRow, c.col, nBody + 1, 2 * c.attempts)
        .setBorder(null, true, null, true, null, null, '#1f4e79', SB.SOLID_THICK);
    });
  }
  sh.getRange(LIST.groupRow, meta.tieCol, nBody + 1, 1)
    .setBorder(null, true, null, true, null, null, '#1f4e79', SB.SOLID_THICK);
}

/**
 * Wywoływane z onEdit: jeśli w edytowanych wierszach listy brakuje reguły (np. skasowany checkbox),
 * przywraca reguły w tych wierszach. Zwraca true, jeśli coś naprawiono.
 */
function repairEditedList_(sh, meta, edited) {
  const r1 = Math.max(edited.getRow(), LIST.firstRow);
  const r2 = Math.min(edited.getLastRow(), LIST.firstRow + LIST.maxRows - 1);
  if (r1 > r2) return false;
  const n = r2 - r1 + 1;
  const dv = sh.getRange(r1, 1, n, listWidth_(meta)).getDataValidations();
  const missing = listRuleSpec_(meta)
    .filter(function (s) { return dv.some(function (row) { return !row[s.col - 1]; }); })
    .map(function (s) { return s.col; });
  if (!missing.length) return false;
  applyListRules_(sh, meta, r1, n, missing);
  return true;
}

/** ⚙️ Setup → Napraw listy startowe. */
function repairStartLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lists = getAllListSheets_(ss);
  repairAllLists_(ss);
  ss.toast('Naprawiono ' + lists.length + ' list(y): checkboxy, reguły, formaty, ramki, numeracja. Dane bez zmian.', '⚙️ Setup', 8);
}

function repairAllLists_(ss) {
  getAllListSheets_(ss).forEach(function (x) {
    applyListRules_(x.sheet, x.meta, LIST.firstRow, LIST.maxRows);
    applyListBorders_(x.sheet, x.meta);
    setNumberingFormula_(x.sheet);
  });
}
