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
 * Je idempotentní — dá se pustit opakovaně.
 *
 * Páruje podle JMÉNA ve sloupci B, ne podle čísla ve sloupci A.
 * Čísla se totiž posunula, když do tabulky přibyli novomanželé,
 * a posunou se zas, až se bude vkládat nebo mazat řádek.
 */
function zapisStoly() {
  var STOLY = {
    'kristyna fizkova':7, 'hana almer':7, 'lubomir almer':7, 'alena pavlouskova':7,
    'pavel pavlousek':9, 'monika mlickova':1, 'ales mlicka':1,
    'tobias mlicka':2, 'natali mlickova':2,
    'zdenek kuffel':1, 'renatka kuffelova':1, 'marie urvalkova':5,
    'dimitra papadopoulu':6, 'vojtech advokat tancos':6, 'ela zichova':6, 'matej habla':6,
    'jarmila lastovicova':2, 'petr lastovica':2,
    'monika kuncarova':5, 'sarlota kuncarova':5,
    'adela tuckova':9, 'tomas musil':9, 'lucie kozakova':10,
    'marcela conova':3, 'david con':3,
    'jakub kozak':8, 'veronika novotna':8, 'vojta kozak':8,
    'nikola havlickova':4, 'martin mikyska':4,
    'martin kuncar':5, 'nikolas kuncar':5,
    'sarka kozakova':8, 'milan kozak':8,
    'andrea stejfova':10, 'marek con':10, 'regina horackova':10,
    'vilem franek':3, 'marketa frankova':3,
    'isabela frankova':4, 'mia frankova':4,
    'tomas brzobohaty':9, 'matej karger':6,
    'honza con':'Novomanželé', 'natali con':'Novomanželé'
  };

  // "Tobias Mlíčka (dítě) - 7" -> "tobias mlicka"
  function klic(jmeno) {
    return String(jmeno)
      .replace(/\(.*?\)/g, ' ')      // závorky pryč
      .replace(/[-–—]\s*\d+\s*$/, '') // koncové " - 7"
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // diakritika pryč
      .toLowerCase().replace(/\s+/g, ' ').trim();
  }

  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('List "' + SHEET_NAME + '" v tabulce není');

  var pocet = sheet.getLastRow() - HEADER_ROW;
  var jmena = sheet.getRange(HEADER_ROW + 1, 2, pocet, 1).getValues();
  var stoly = sheet.getRange(HEADER_ROW + 1, COL_STUL, pocet, 1).getValues();

  var zapsano = 0, nenalezeni = [];
  for (var i = 0; i < jmena.length; i++) {
    var j = String(jmena[i][0]).trim();
    if (!j) continue;
    var k = klic(j);
    if (STOLY.hasOwnProperty(k)) { stoly[i][0] = STOLY[k]; zapsano++; }
    else nenalezeni.push(j);
  }

  sheet.getRange(HEADER_ROW + 1, COL_STUL, pocet, 1).setValues(stoly);
  Logger.log('Zapsáno stolů: ' + zapsano + ' z ' + Object.keys(STOLY).length + ' očekávaných');
  if (nenalezeni.length) Logger.log('V seznamu stolů chybí: ' + nenalezeni.join(', '));
}
