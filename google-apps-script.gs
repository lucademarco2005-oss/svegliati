/**
 * SVEGLIATI — salva i contatti della landing page sul Google Foglio.
 *
 * COME SI INSTALLA (una volta sola):
 * 1. Apri il foglio "SVEGLIATI — Contatti dalla locandina".
 * 2. Menu Estensioni > Apps Script.
 * 3. Cancella tutto il codice di esempio e incolla questo file.
 * 4. In alto clicca Distribuisci > Nuova distribuzione > tipo "App web".
 *    - Esegui come: Me
 *    - Chi ha accesso: Chiunque
 * 5. Autorizza quando te lo chiede, poi copia l'URL che finisce con /exec.
 * 6. Incolla quell'URL dentro index.html, nella riga  var ENDPOINT = "";
 */

var SHEET_ID = '1g0UlzP4eKGTaf0xu3s33nsldMIYc_wUaVj6IGoZjeZw';
var NOTIFICA_EMAIL = ''; // opzionale: metti un'email per ricevere anche un avviso a ogni contatto

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Data', 'Nome e cognome', 'Telefono/WhatsApp', 'Email', 'Interesse', 'Privacy', 'Fonte']);
    }

    sheet.appendRow([
      new Date(),
      d.nome || '',
      "'" + (d.telefono || ''), // apice iniziale: cosi' il numero non perde lo zero
      d.email || '',
      d.interesse || '',
      d.privacy || '',
      d.fonte || ''
    ]);

    if (NOTIFICA_EMAIL) {
      MailApp.sendEmail(
        NOTIFICA_EMAIL,
        'Nuovo contatto SVEGLIATI: ' + (d.nome || ''),
        'Nome: ' + (d.nome || '') +
        '\nTelefono: ' + (d.telefono || '') +
        '\nEmail: ' + (d.email || '') +
        '\nInteresse: ' + (d.interesse || '')
      );
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('SVEGLIATI endpoint attivo');
}
