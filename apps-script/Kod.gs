/**
 * Zápis výběru jídla z thecon.cz/jidlo do master tabulky hostů.
 *
 * Nasazení (jednorázově):
 *   1. Vložit tento kód do Apps Scriptu, uložit
 *   2. Nahoře vybrat funkci "test" a dát Spustit — ověří přístup k tabulce
 *      a vyžádá si oprávnění. Musí vypsat počet hostů.
 *   3. Implementovat → Nové nasazení → typ "Webová aplikace"
 *        Spustit jako: Já
 *        Kdo má přístup: Kdokoli
 *   4. Zkopírovat URL webové aplikace a předat ji do jidlo/index.html
 *
 * Po každé změně kódu je nutné nasadit znovu (Nasadit → Spravovat nasazení → tužka → Nová verze),
 * jinak běží pořád stará verze.
 */

// Tabulku otevíráme podle ID, ne přes getActive() — díky tomu je jedno,
// jestli je skript navázaný na tabulku, nebo je to samostatný projekt.
var SPREADSHEET_ID = '1NTbHAcz9_QeUmkxDOvuYsiH6JTRa4D9jFySBMehBbvs';
var SHEET_NAME = 'Master – hosté';
var HEADER_ROW = 4;      // řádek s názvy sloupců
var COL_ID = 1;          // A = #
var COL_PREDKRM = 18;    // R
var COL_POLEVKA = 19;    // S
var COL_VYPLNENO = 20;   // T
var COL_STUL     = 9;    // I — zasedací pořádek

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);   // dva hosté odesílající naráz si nepřepíšou řádky

  try {
    var payload = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('List "' + SHEET_NAME + '" v tabulce není');

    // Mapa id (#) → číslo řádku
    var ids = sheet.getRange(HEADER_ROW + 1, COL_ID, sheet.getLastRow() - HEADER_ROW, 1).getValues();
    var rowById = {};
    ids.forEach(function (r, i) {
      if (r[0] !== '' && r[0] !== null) rowById[String(r[0]).trim()] = HEADER_ROW + 1 + i;
    });

    var stamp = Utilities.formatDate(new Date(), 'Europe/Prague', 'd.M.yyyy H:mm');
    var zapsano = [];
    var nenalezeno = [];

    (payload.osoby || []).forEach(function (o) {
      var row = rowById[String(o.id).trim()];
      if (!row) { nenalezeno.push(o.id); return; }
      sheet.getRange(row, COL_PREDKRM).setValue(o.predkrm);
      sheet.getRange(row, COL_POLEVKA).setValue(o.polevka);
      sheet.getRange(row, COL_VYPLNENO).setValue(stamp);
      zapsano.push(o.id);
    });

    // Poznámka od hosta se připíše k tomu, kdo formulář odeslal — stávající text nepřepisujeme
    if (payload.poznamka && payload.osoby && payload.osoby.length) {
      var prvni = rowById[String(payload.osoby[0].id).trim()];
      if (prvni) {
        var cell = sheet.getRange(prvni, 17);   // Q = Poznámka
        var stary = String(cell.getValue() || '').trim();
        var novy = 'K jídlu: ' + payload.poznamka;
        cell.setValue(stary ? stary + ' | ' + novy : novy);
      }
    }

    return odpoved({ ok: true, zapsano: zapsano, nenalezeno: nenalezeno });
  } catch (err) {
    return odpoved({ ok: false, chyba: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return odpoved({ ok: true, info: 'Endpoint pro výběr jídla běží.' });
}

/** Spusť ručně v editoru — ověří, že skript vidí tabulku a má oprávnění. */
function test() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('List "' + SHEET_NAME + '" v tabulce není');
  var pocet = sheet.getLastRow() - HEADER_ROW;
  var hlavicka = sheet.getRange(HEADER_ROW, COL_PREDKRM, 1, 3).getValues()[0];
  Logger.log('Tabulka: ' + sheet.getParent().getName());
  Logger.log('Řádků s hosty: ' + pocet);
  Logger.log('Sloupce R–T: ' + hlavicka.join(' | '));
}

function odpoved(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Jednorázový zápis zasedacího pořádku do sloupce I (Stůl).
 * Spusť ručně v editoru: nahoře vyber "zapisStoly" a dej Spustit.
 * Je idempotentní — dá se pustit opakovaně, jen přepíše stejné hodnoty.
 */
function zapisStoly() {
  // id hosta (sloupec A) -> číslo stolu
  var STOLY = {
    1:7,  2:7,  3:7,  4:7,  5:9,  6:1,  7:1,  8:2,  9:2,  10:1,
    11:1, 12:5, 13:6, 14:6, 15:6, 16:6, 17:2, 18:2, 19:5, 20:5,
    21:9, 22:9, 23:10, 24:3, 25:3, 26:8, 27:8, 28:8, 29:4, 30:4,
    31:5, 32:5, 33:8, 34:8, 35:10, 36:10, 37:10, 38:3, 39:3, 40:4,
    41:4, 42:9, 43:6
  };

  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('List "' + SHEET_NAME + '" v tabulce není');

  var pocetRadku = sheet.getLastRow() - HEADER_ROW;
  var ids = sheet.getRange(HEADER_ROW + 1, COL_ID, pocetRadku, 1).getValues();
  var stavajici = sheet.getRange(HEADER_ROW + 1, COL_STUL, pocetRadku, 1).getValues();

  var zapsano = 0, bezStolu = [];
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i][0]).trim();
    if (id === '' || id === 'null') continue;
    if (STOLY.hasOwnProperty(id)) {
      stavajici[i][0] = STOLY[id];
      zapsano++;
    } else {
      bezStolu.push(id);
    }
  }

  sheet.getRange(HEADER_ROW + 1, COL_STUL, pocetRadku, 1).setValues(stavajici);
  Logger.log('Zapsáno stolů: ' + zapsano);
  if (bezStolu.length) Logger.log('Bez přiřazeného stolu (id): ' + bezStolu.join(', '));
}
