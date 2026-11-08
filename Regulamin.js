/**
 * Zakładka „Regulamin” i wydruk do PDF.
 *
 * Zakładka: kolumna A = styl (lista, puste = rozpoznawany automatycznie), kolumna B = treść, wiersz = akapit.
 * „Regulamin” to zakładka z danymi - setupTemplate() zakłada ją tylko, jeśli nie istnieje.
 */

const SHEET_RULES = 'Regulamin';
const RULES = { headerRow: 1, firstRow: 2 };

/** Style akapitów: klucz = wartość w kolumnie Styl. */
const RULE_STYLES = ['tytuł', 'podtytuł', 'sekcja', 'podsekcja', 'akapit', 'punkt', 'numer', 'link', 'wyróżnienie'];

const DEFAULT_RULES_TEXT = `
REGULAMIN
II Charytatywne Zawody Mikołajkowe Kettlebell Hardstyle – Best Move
1. CEL ZAWODÓW
Zawody mają charakter charytatywny. Całość wpisowego zostanie przekazana na leczenie i rehabilitację Hugona Szatrowskiego.
https://www.fundacjaavalon.pl/nas.../hugon_szatrowski_19577/
Celem wydarzenia jest również promowanie Kettlebell Hardstyle, zdrowej rywalizacji i wspólnego pomagania.
2. DATA I MIEJSCE
6 grudnia 2026 r. (niedziela)
Sala treningowa Best Move, Brzeg Dolny.
3. KONKURENCJE
1. One Arm Press – 1RM – wyciskanie jednorącz, jedno maksymalne powtórzenie.
2. Turkish Get Up (TGU) – 1RM – tureckie wstawanie, jedno maksymalne powtórzenie.
4. KATEGORIE I LIMIT UCZESTNIKÓW
Rywalizacja odbywa się w trzech kategoriach:
- Kobiety
- Mężczyźni
- Juniorzy w wieku 10–16 lat
W kategoriach kobiet i mężczyzn nie obowiązuje podział na klasy wagowe.
Juniorzy startują wyłącznie w konkurencji Turkish Get Up – 1RM.
Obowiązuje limit 40 uczestników, w tym:
- 30 miejsc dla zawodników dorosłych.
- 10 miejsc dla juniorów w wieku 10–16 lat.
O przyjęciu decyduje kolejność prawidłowo dokonanych zgłoszeń.
Start zawodników niepełnoletnich wymaga pisemnej zgody rodzica lub opiekuna prawnego albo obecności rodzica lub opiekuna prawnego podczas zawodów.
5. WYNIK I KLASYFIKACJA
Każdy zawodnik zostaje zważony przed rozpoczęciem rywalizacji.
Wyniki będą obliczane według przelicznika wagowego opracowanego przez Kacpra Bajdę, z powodzeniem stosowanego podczas Mistrzostw Polski Kettlebell Hardstyle w Bochni.
Przelicznik obowiązuje we wszystkich konkurencjach i kategoriach – kobiet, mężczyzn oraz juniorów.
O miejscu w klasyfikacji decyduje najwyższy wynik uzyskany po zastosowaniu przelicznika.
Każdy zawodnik ma trzy podejścia w każdej konkurencji, w której startuje.
Po nieudanej próbie zawodnik może ponowić podejście z tym samym ciężarem lub wybrać cięższy odważnik. Nie może zmniejszyć ciężaru w kolejnej próbie.
Do klasyfikacji zaliczany jest najlepszy prawidłowo wykonany bój.
6. ZASADY TECHNICZNE – HARDSTYLE
W obu konkurencjach oceniana jest poprawność techniczna, stabilizacja i pełna kontrola nad odważnikiem.
A. ONE ARM PRESS – 1RM
1. Zawodnik samodzielnie zarzuca odważnik do pozycji rack. Pięść ręki pracującej znajduje się poniżej brody.
2. Ruch wyciskania rozpoczyna się na komendę sędziego.
3. Odważnik musi zostać wyciśnięty siłą ramienia, bez pomocy nóg.
4. Podczas wyciskania kolana pozostają wyprostowane. Niedozwolone jest uginanie kolan, wybicie z nóg oraz wykonywanie push press lub jerk.
5. Całe stopy muszą pozostawać w kontakcie z podłożem.
6. Nadgarstek ręki pracującej pozostaje prosty, w jednej linii z przedramieniem.
7. W pozycji końcowej wymagany jest pełny wyprost łokcia i stabilne utrzymanie odważnika nad głową.
8. Zawodnik musi pokazać kontrolę nad ciężarem, bez natychmiastowego opuszczania odważnika.
9. Po zaliczeniu próby odważnik należy sprowadzić do pozycji rack i kontrolowanym ruchem odstawić na podłogę.
10. Utrata równowagi, pomoc nóg, brak pełnego wyprostu lub brak kontroli nad ciężarem skutkują niezaliczeniem próby.
B. TURKISH GET UP – 1RM
1. Zawodnik rozpoczyna ćwiczenie w pozycji leżącej na plecach i samodzielnie wyciska odważnik do wyprostu ramienia.
2. Od momentu wyciśnięcia odważnika łokieć ręki pracującej musi pozostawać zablokowany przez cały czas trwania ćwiczenia.
3. Nadgarstek ręki trzymającej kettla musi pozostawać prosty, w jednej linii z przedramieniem.
4. Zawodnik wykonuje pełne tureckie wstawanie, przechodząc z pozycji leżącej do stojącej, a następnie wraca do pozycji leżącej.
5. W pozycji stojącej zawodnik musi zatrzymać się na minimum 1 sekundę, utrzymując wyprostowane ramię z odważnikiem nad głową i głowę skierowaną do przodu.
6. W pozycji stojącej stopy nie muszą znajdować się równolegle. Dopuszczalne jest ustawienie jednej stopy przed drugą, pod warunkiem że pięta przedniej stopy nie znajduje się dalej niż palce tylnej.
7. Po zatrzymaniu w pozycji stojącej zawodnik rozpoczyna kontrolowane zejście do pozycji leżącej.
8. Po powrocie do pozycji leżącej należy ponownie pokazać kontrolę nad ciężarem, utrzymując wyprostowane ramię przez minimum 1 sekundę.
9. Odważnik należy bezpiecznie i kontrolowanie odłożyć na podłogę. Zabronione jest jego odrzucanie.
10. Ugięcie łokcia ręki pracującej, utrata kontroli nad odważnikiem, brak wymaganej pauzy lub nieukończenie całego ruchu skutkują niezaliczeniem próby.
DODATKOWE ZASADY:
- Rękę pracującą zawodnik wybiera samodzielnie.
- Próby rozpoczynają się na komendę sędziego.
- Sędzia ocenia prawidłowość wykonania całego boju.
- Decyzja o zaliczeniu lub niezaliczeniu próby należy do sędziego stanowiska.
7. SPRZĘT I ZASADY OGÓLNE
- Start odbywa się boso lub w skarpetkach, bez obuwia.
- Zabronione jest korzystanie z pasów, pasków, rękawiczek, opasek i bandaży wspomagających wykonanie boju.
- Dozwolone jest stosowanie magnezji.
- Wszystkie ćwiczenia wykonywane są w formule RAW, bez sprzętu wspomagającego.
- Zawodnik musi samodzielnie i bezpiecznie operować odważnikiem.
- Organizator i sędzia mają prawo przerwać próbę, jeżeli jej dalsze wykonywanie zagraża bezpieczeństwu zawodnika lub innych osób.
8. WPISOWE I ZAPISY
Wpisowe wynosi minimum 100 zł.
Całość wpłaty trafia bezpośrednio na zbiórkę Hugona Szatrowskiego:
https://www.fundacjaavalon.pl/nas.../hugon_szatrowski_19577/
Aby zapisać się na zawody, należy:
1. Wpłacić cegiełkę w wysokości minimum 100 zł na zbiórkę Hugona.
2. Przesłać organizatorowi przez Messenger wiadomość zawierającą imię, nazwisko oraz potwierdzenie wpłaty.
Po potwierdzeniu zgłoszenia zawodnik zostaje wpisany na listę startową.
Liczba miejsc jest ograniczona do 40 osób, w tym 10 juniorów.
Zapisy trwają od 03.11 do 30.11 2026
9. NAGRODY
Każdy uczestnik otrzyma pamiątkowy medal oraz pakiet startowy.
Dla najlepszych zawodników przewidziane są puchary i nagrody.
10. KONTAKT
W przypadku pytań dotyczących regulaminu, konkurencji lub zapisów proszę o kontakt przez Messenger.
Dźwigamy, pomagamy i dobrze się bawimy!
`;

/** Rozpoznaje styl akapitu. index = numer niepustego akapitu (0 = pierwszy). */
function detectRuleStyle_(text, index) {
  const t = String(text).trim();
  if (index === 0) return 'tytuł';
  if (index === 1) return 'podtytuł';
  if (/^https?:\/\//i.test(t)) return 'link';
  const hasLetters = /[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]/.test(t);
  const isUpper = hasLetters && t === t.toUpperCase();
  if (/^\d+\.\s/.test(t) && isUpper) return 'sekcja';
  if (/^[A-ZĄĆĘŁŃÓŚŹŻ]\.\s/.test(t) && isUpper) return 'podsekcja';
  if (isUpper && /:$/.test(t)) return 'podsekcja';
  if (/^[-–•]\s/.test(t)) return 'punkt';
  if (/^\d+\.\s/.test(t)) return 'numer';
  return 'akapit';
}

/** Zamienia wiersze [styl, tekst] na akapity {style, text} - puste style rozpoznawane automatycznie. */
function parseRules_(rows) {
  const out = [];
  rows.forEach(function (r) {
    const text = String(r[1]).trim();
    if (!text) return;
    const style = String(r[0]).trim().toLowerCase();
    out.push({ style: RULE_STYLES.indexOf(style) >= 0 ? style : detectRuleStyle_(text, out.length), text: text });
  });
  return out;
}

// ───────────────────────── Zakładka ─────────────────────────

function buildRulesSheet_(ss) {
  const sh = ss.insertSheet(SHEET_RULES, ss.getSheets().length);
  const lines = DEFAULT_RULES_TEXT.trim().split('\n');
  const parsed = parseRules_(lines.map(function (l) { return ['', l]; }));
  parsed[parsed.length - 1].style = 'wyróżnienie';

  sh.getRange(RULES.headerRow, 1, 1, 2)
    .setValues([['Styl (puste = automatycznie)', 'Treść regulaminu - wiersz = akapit. Wydruk: 🏋️ Zawody → Drukuj regulamin (PDF)']])
    .setFontWeight('bold').setBackground('#d9e2f3').setWrap(true).setVerticalAlignment('middle');

  const n = parsed.length;
  sh.getRange(RULES.firstRow, 1, n, 2).setValues(parsed.map(function (p) { return [p.style, p.text]; }));

  const maxRows = Math.max(n + 50, 200);
  if (sh.getMaxRows() < maxRows + 1) sh.insertRowsAfter(sh.getMaxRows(), maxRows + 1 - sh.getMaxRows());
  if (sh.getMaxColumns() > 2) sh.deleteColumns(3, sh.getMaxColumns() - 2);

  sh.getRange(RULES.firstRow, 1, maxRows, 1)
    .setFontColor('#666666')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(RULE_STYLES, true).setAllowInvalid(false)
      .setHelpText('Styl akapitu w PDF. Puste = rozpoznawany automatycznie.').build());
  sh.getRange(RULES.firstRow, 2, maxRows, 1).setWrap(true).setVerticalAlignment('top');

  // Podgląd stylów w arkuszu (tylko wizualnie - w PDF decyduje kolumna Styl).
  parsed.forEach(function (p, i) {
    const cell = sh.getRange(RULES.firstRow + i, 2);
    if (p.style === 'tytuł') cell.setFontSize(16).setFontWeight('bold');
    if (p.style === 'podtytuł') cell.setFontSize(12).setFontWeight('bold');
    if (p.style === 'sekcja') cell.setFontWeight('bold').setBackground('#eef3fb');
    if (p.style === 'podsekcja' || p.style === 'wyróżnienie') cell.setFontWeight('bold');
    if (p.style === 'link') cell.setFontColor('#1a5fb4');
  });

  sh.setColumnWidth(1, 120);
  sh.setColumnWidth(2, 820);
  sh.setFrozenRows(RULES.headerRow);
  return sh;
}

// ───────────────────────── PDF ─────────────────────────

function printRules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_RULES);
  if (!sh) throw new Error('Brak zakładki „' + SHEET_RULES + '” - uruchom ⚙️ Setup → Zbuduj szablon.');

  const last = sh.getLastRow();
  const rows = last >= RULES.firstRow ? sh.getRange(RULES.firstRow, 1, last - RULES.firstRow + 1, 2).getValues() : [];
  const paragraphs = parseRules_(rows);
  if (!paragraphs.length) throw new Error('Zakładka „' + SHEET_RULES + '” jest pusta.');

  const title = paragraphs[0].text;
  const fileName = ('Regulamin - ' + (paragraphs[1] ? paragraphs[1].text : title))
    .replace(/[\\/:*?"<>|]/g, '-').slice(0, 120) + '.pdf';

  const html = rulesHtml_(paragraphs);
  const pdf = Utilities.newBlob(html, 'text/html', 'regulamin.html').getAs('application/pdf').setName(fileName);
  const b64 = Utilities.base64Encode(pdf.getBytes());

  const dialog = HtmlService.createHtmlOutput(downloadDialogHtml_(b64, fileName)).setWidth(420).setHeight(170);
  SpreadsheetApp.getUi().showModalDialog(dialog, '🖨️ Regulamin (PDF)');
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** HTML regulaminu do konwersji na PDF (A4). */
function rulesHtml_(paragraphs) {
  const accent = '#1f4e79';
  const css = [
    '@page { size: A4; margin: 16mm 18mm; }',
    'body { font-family: Arial, Helvetica, sans-serif; font-size: 10.5pt; line-height: 1.45; color: #222; margin: 0; }',
    '.tytul { font-size: 24pt; font-weight: bold; text-align: center; letter-spacing: 3pt; color: ' + accent + '; margin: 0 0 4pt; }',
    '.podtytul { font-size: 13pt; font-weight: bold; text-align: center; color: #333; margin: 0 0 16pt; padding-bottom: 10pt; border-bottom: 2pt solid ' + accent + '; }',
    '.sekcja { font-size: 12pt; font-weight: bold; color: ' + accent + '; margin: 14pt 0 5pt; padding-bottom: 2pt; border-bottom: 0.75pt solid #c0c7d4; page-break-after: avoid; }',
    '.podsekcja { font-size: 11pt; font-weight: bold; color: #333; margin: 10pt 0 4pt; page-break-after: avoid; }',
    '.akapit { margin: 0 0 4pt; text-align: justify; }',
    '.punkt { margin: 0 0 2pt 14pt; padding-left: 10pt; text-indent: -10pt; }',
    '.numer { margin: 0 0 3pt 14pt; padding-left: 16pt; text-indent: -16pt; }',
    '.link { margin: 0 0 4pt; color: #1a5fb4; word-wrap: break-word; }',
    '.link a { color: #1a5fb4; text-decoration: none; }',
    '.wyroznienie { font-size: 12pt; font-weight: bold; text-align: center; color: ' + accent + '; margin: 18pt 0 0; }',
  ].join('\n');

  const body = paragraphs.map(function (p) {
    const t = escapeHtml_(p.text);
    switch (p.style) {
      case 'tytuł': return '<div class="tytul">' + t + '</div>';
      case 'podtytuł': return '<div class="podtytul">' + t + '</div>';
      case 'sekcja': return '<div class="sekcja">' + t + '</div>';
      case 'podsekcja': return '<div class="podsekcja">' + t + '</div>';
      case 'punkt': return '<div class="punkt">•&nbsp;&nbsp;' + escapeHtml_(p.text.replace(/^[-–•]\s*/, '')) + '</div>';
      case 'numer': return '<div class="numer">' + t + '</div>';
      case 'link': return '<div class="link"><a href="' + t + '">' + t + '</a></div>';
      case 'wyróżnienie': return '<div class="wyroznienie">' + t + '</div>';
      default: return '<div class="akapit">' + t + '</div>';
    }
  }).join('\n');

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' + css + '</style></head><body>' + body + '</body></html>';
}

/** Okienko z automatycznym pobraniem PDF (i przyciskiem na wypadek blokady). */
function downloadDialogHtml_(b64, fileName) {
  const name = escapeHtml_(fileName);
  return '<div style="font-family:Arial,sans-serif;font-size:14px">' +
    '<p>Regulamin gotowy: <b>' + name + '</b></p>' +
    '<p><a id="dl" download="' + name + '" href="data:application/pdf;base64,' + b64 + '" ' +
    'style="display:inline-block;padding:8px 16px;background:#1f4e79;color:#fff;border-radius:4px;text-decoration:none">⬇️ Pobierz PDF</a></p>' +
    '</div>' +
    '<script>try { document.getElementById("dl").click(); } catch (e) {}</script>';
}
