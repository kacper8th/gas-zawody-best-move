/**
 * Kontrola danych zawodników względem stanu z chwili blokady („wzorzec”) i przywracanie.
 *
 * Wzorzec (snapshot) = imię, nazwisko, waga w kolejności z listy, zapisany przy 🔒 Zablokuj dane zawodników.
 * Zawodnicy są dopasowywani po imieniu i nazwisku, nie po numerze wiersza - dzięki temu sortowanie
 * i usunięte wiersze są wykrywane, a przywracanie przenosi całe wiersze razem z próbami.
 */

/** Klucz zawodnika: imię + nazwisko bez wielkości liter i zbędnych spacji ('' = pusty wiersz). */
function athleteKey_(first, last) {
  return (String(first) + ' ' + String(last)).toLowerCase().replace(/\s+/g, ' ').trim();
}

/**
 * Klucze z numerem wystąpienia: ten sam zawodnik wpisany dwa razy to „ala kot#1” i „ala kot#2”,
 * żeby duplikaty nie myliły dopasowania. Pusty wiersz = ''.
 */
function occurrenceKeys_(rows) {
  const seen = {};
  return rows.map(function (r) {
    const k = athleteKey_(r[0], r[1]);
    if (!k) return '';
    seen[k] = (seen[k] || 0) + 1;
    return k + '#' + seen[k];
  });
}

/** Wzorzec listy z właściwości dokumentu albo null. */
function readAthletesSnapshot_(sheet) {
  const raw = PropertiesService.getDocumentProperties().getProperty(athletesSnapKey_(sheet));
  return raw ? JSON.parse(raw) : null;
}

/**
 * Czysta funkcja: różnice między wzorcem a bieżącą listą.
 * snapCells / curRows: [[imię, nazwisko, waga], ...] w kolejności wierszy.
 */
function lockDiff_(snapCells, curRows) {
  const label = function (r) { return (String(r[0]).trim() + ' ' + String(r[1]).trim()).trim(); };
  const snap = [], cur = [];
  const snapKeys = occurrenceKeys_(snapCells), curKeys = occurrenceKeys_(curRows);
  snapCells.forEach(function (r, i) { if (snapKeys[i]) snap.push({ key: snapKeys[i], pos: i, row: r }); });
  curRows.forEach(function (r, i) { if (curKeys[i]) cur.push({ key: curKeys[i], pos: i, row: r }); });

  const curByKey = {};
  cur.forEach(function (c) { if (!curByKey[c.key]) curByKey[c.key] = c; });
  const snapByKey = {};
  snap.forEach(function (s) { if (!snapByKey[s.key]) snapByKey[s.key] = s; });

  const missing = [], added = [], weight = [];
  snap.forEach(function (s) {
    const c = curByKey[s.key];
    if (!c) { missing.push({ name: label(s.row), nr: s.pos + 1 }); return; }
    const a = toNum_(s.row[2]), b = toNum_(c.row[2]);
    if (a !== b) weight.push({ name: label(s.row), from: a, to: b });
  });
  cur.forEach(function (c) { if (!snapByKey[c.key]) added.push({ name: label(c.row), row: c.pos + 1 }); });

  // Kolejność: porównanie względnej kolejności zawodników obecnych w obu (usunięcie wiersza jej nie zmienia).
  const snapOrder = snap.filter(function (s) { return curByKey[s.key]; }).map(function (s) { return s.key; });
  const curOrder = cur.filter(function (c) { return snapByKey[c.key]; }).map(function (c) { return c.key; });
  let moved = null;
  for (let i = 0; i < snapOrder.length; i++) {
    if (snapOrder[i] !== curOrder[i]) { moved = { name: label(snapByKey[snapOrder[i]].row), nr: snapByKey[snapOrder[i]].pos + 1 }; break; }
  }

  return { missing: missing, added: added, weight: weight, moved: moved,
    any: !!(missing.length || added.length || weight.length || moved) };
}

/** Czysta funkcja: komunikaty dla różnic jednej listy. */
function lockDiffMessages_(category, d) {
  const fmt = function (v) { return v === null ? 'brak' : String(v).replace('.', ','); };
  const out = [];
  if (d.moved) out.push(category + ': zmieniona kolejność zawodników (np. ' + d.moved.name + ' był(a) na pozycji ' + d.moved.nr + ')');
  d.weight.forEach(function (w) { out.push(category + ': ' + w.name + ' - waga ' + fmt(w.from) + ' → ' + fmt(w.to) + ' mimo blokady'); });
  d.missing.forEach(function (m) { out.push(category + ': brak zawodnika ' + m.name + ' (był nr ' + m.nr + ')'); });
  d.added.forEach(function (a) { out.push(category + ': dodany zawodnik ' + a.name + ' mimo blokady'); });
  return out;
}

/** Różnice wszystkich list względem wzorców (tylko przy włączonej blokadzie). */
function collectLockIssues_(ss) {
  if (!getLockState_().athletes) return [];
  const out = [];
  getListSheets_(ss).forEach(function (x) {
    const snap = readAthletesSnapshot_(x.sheet);
    if (!snap) return;
    const a = athletesArea_();
    const cur = x.sheet.getRange(a.row, a.col, a.rows, a.cols).getValues();
    lockDiffMessages_(x.meta.category, lockDiff_(snap.cells, cur)).forEach(function (m) { out.push(m); });
  });
  return out;
}

/** 🏋️ Zawody → Przywróć dane zawodników z blokady. */
function restoreAthletesFromLock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (!getLockState_().athletes) {
    ui.alert('Przywróć dane zawodników', 'Blokada danych zawodników nie jest włączona - nie ma stanu do przywrócenia.', ui.ButtonSet.OK);
    return;
  }

  const plans = [];
  const messages = [];
  getListSheets_(ss).forEach(function (x) {
    const snap = readAthletesSnapshot_(x.sheet);
    if (!snap) return;
    const width = listWidth_(x.meta) - 1; // od Imienia do Uwag/Dogrywki (bez Nr - jest formułą)
    const full = x.sheet.getRange(LIST.firstRow, 2, LIST.maxRows, width).getValues();
    const d = lockDiff_(snap.cells, full.map(function (r) { return r.slice(0, 3); }));
    if (!d.any) return;
    lockDiffMessages_(x.meta.category, d).forEach(function (m) { messages.push(m); });
    plans.push({ sheet: x.sheet, category: x.meta.category, plan: planRestore_(snap.cells, full, LIST.maxRows) });
  });

  if (!plans.length) {
    ui.alert('Przywróć dane zawodników', '✅ Dane zawodników są zgodne ze stanem z blokady.', ui.ButtonSet.OK);
    return;
  }
  const ans = ui.alert('Przywróć dane zawodników',
    'Różnice względem stanu z blokady:\n\n• ' + messages.slice(0, 30).join('\n• ') +
    '\n\nPrzywrócić? Próby zostają przy swoich zawodnikach.', ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;

  const summary = [];
  plans.forEach(function (p) {
    const range = p.sheet.getRange(LIST.firstRow, 2, p.plan.rows.length, p.plan.rows[0].length);
    range.setValues(p.plan.rows);
    range.setBackground(null); // stare podświetlenia walidacji nie pasują już do wierszy
    p.plan.notes.forEach(function (n) { summary.push(p.category + ': ' + n); });
  });
  rebuildResults();
  ui.alert('Przywróć dane zawodników',
    '✅ Przywrócono kolejność i dane zawodników.' + (summary.length ? '\n\nDo sprawdzenia:\n• ' + summary.join('\n• ') : ''),
    ui.ButtonSet.OK);
}

/**
 * Czysta funkcja: nowy układ wierszy listy wg wzorca.
 * snapCells: [[imię, nazwisko, waga]]; full: bieżące wiersze od kolumny Imię do Dogrywka; maxRows: wysokość listy.
 * Zwraca { rows, notes }. Wiersze bez zawodnika, ale z wpisanymi danymi, i zawodnicy dodani - na koniec (nic nie ginie).
 */
function planRestore_(snapCells, full, maxRows) {
  const width = full[0].length;
  const blank = function () { return new Array(width).fill(''); };
  const used = {};
  const byKey = {};
  occurrenceKeys_(full).forEach(function (k, i) { if (k) byKey[k] = i; });
  const snapKeys = occurrenceKeys_(snapCells);

  const rows = [], notes = [];
  snapCells.forEach(function (s, si) {
    const k = snapKeys[si];
    if (!k) { rows.push(blank()); return; }
    const i = byKey[k];
    if (i === undefined) {
      const r = blank();
      r[0] = s[0]; r[1] = s[1]; r[2] = s[2];
      rows.push(r);
      notes.push((String(s[0]) + ' ' + String(s[1])).trim() + ' - przywrócony bez prób, próby odzyskaj z historii wersji (Plik → Historia wersji)');
      return;
    }
    used[i] = true;
    const r = full[i].slice();
    r[2] = s[2]; // waga ze wzorca
    rows.push(r);
  });

  full.forEach(function (r, i) {
    if (used[i]) return;
    const hasData = r.some(function (v) { return v !== '' && v !== false && v !== null; });
    if (!hasData) return;
    rows.push(r.slice());
    const name = (String(r[0]) + ' ' + String(r[1])).trim();
    notes.push(name ? name + ' - dodany po blokadzie, przeniesiony na koniec listy (usuń, jeśli niepotrzebny)'
      : 'wiersz bez zawodnika z wpisanymi próbami - przeniesiony na koniec listy');
  });

  while (rows.length < maxRows) rows.push(blank());
  return { rows: rows.slice(0, Math.max(maxRows, rows.length)), notes: notes };
}

/**
 * Cofnięcie ręcznej edycji danych zawodników przy włączonej blokadzie (wywoływane z onEdit).
 * Pojedyncza komórka - dokładnie poprzednia wartość. Blok - tylko gdy kolejność listy zgadza się ze wzorcem.
 * Zwraca 'reverted', 'skipped' (lista różni się od wzorca) albo null (edycja poza danymi zawodników).
 */
function revertAthletesEdit_(sh, e) {
  const a = athletesArea_();
  const box = intersect_(
    { r1: e.range.getRow(), r2: e.range.getLastRow(), c1: e.range.getColumn(), c2: e.range.getLastColumn() },
    { r1: a.row, r2: a.row + a.rows - 1, c1: a.col, c2: a.col + a.cols - 1 });
  if (!box) return null;

  if (e.range.getNumRows() === 1 && e.range.getNumColumns() === 1) {
    let v = e.oldValue === undefined ? '' : e.oldValue;
    if (typeof v === 'string' && /^-?\d+(\.\d+)?$/.test(v)) v = Number(v);
    e.range.setValue(v);
    return 'reverted';
  }

  const snap = readAthletesSnapshot_(sh);
  if (!snap) return null;
  const cur = sh.getRange(a.row, a.col, a.rows, a.cols).getValues();
  if (!positionsConsistent_(snap.cells, cur, box.r1 - a.row, box.r2 - a.row)) return 'skipped';
  restoreFromSnapshot_(sh, athletesSnapKey_(sh), e.range);
  return 'reverted';
}

/** Czysta funkcja: czy poza wierszami [ex1, ex2] (indeksy od 0) zawodnicy stoją na tych samych pozycjach co we wzorcu. */
function positionsConsistent_(snapCells, curRows, ex1, ex2) {
  const n = Math.max(snapCells.length, curRows.length);
  for (let i = 0; i < n; i++) {
    if (i >= ex1 && i <= ex2) continue;
    const s = snapCells[i] || ['', ''];
    const c = curRows[i] || ['', ''];
    if (athleteKey_(s[0], s[1]) !== athleteKey_(c[0], c[1])) return false;
  }
  return true;
}
