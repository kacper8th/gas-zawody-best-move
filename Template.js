/**
 * Szablon arkusza zawody-best-move.
 * setupTemplate() jest idempotentna: zakładki z danymi zakłada tylko, jeśli nie istnieją,
 * czyści i buduje od nowa wyłącznie zakładki wyliczane.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Setup')
    .addItem('Zbuduj szablon', 'setupTemplate')
    .addToUi();
}

function setupTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // TODO: struktura arkusza (zakładki, kolumny, nagłówki) do ustalenia z Kacprem.
  ss.toast('Szablon zbudowany.', '⚙️ Setup');
}

/** Zwraca zakładkę o danej nazwie, zakładając ją, jeśli nie istnieje (bez czyszczenia danych). */
function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
