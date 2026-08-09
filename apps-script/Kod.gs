/**
 * Zápis výběru jídla z thecon.cz/jidlo do master tabulky hostů.
 *
 * Nasazení (jednorázově):
 *   1. V tabulce: Rozšíření → Apps Script
 *   2. Vložit tento kód, uložit
 *   3. Nasadit → Nové nasazení → typ "Webová aplikace"
 *        Spustit jako: Já
 *        Kdo má přístup: Kdokoli
 *   4. Zkopírovat URL webové aplikace a předat ji do jidlo/index.html
 *
 * Po každé změně kódu je nutné nasadit znovu (Nasadit → Spravovat nasazení → tužka → Nová verze),
 * jinak běží pořád stará verze.
 */

var SHEET_NAME = 'Master – hosté';
var HEADER_ROW = 4;      // řádek s názvy sloupců
var COL_ID = 1;          // A = #
var COL_PREDKRM = 18;    // R
var COL_POLEVKA = 19;    // S
var COL_VYPLNENO = 20;   // T

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);   // dva hosté odesílající naráz si nepřepíšou řádky

  try {
    var payload = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
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

function odpoved(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
