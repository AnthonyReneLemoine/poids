/**
 * Outils de migration interne vers une architecture 100% Google Apps Script + Google Sheets.
 *
 * Ce script ne dépend d'aucun service externe.
 */

function migrerHistoriqueVersSauvegarde() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backup = getBackupSheet();
  const types = ['human', 'cat'];
  let count = 0;

  for (const type of types) {
    const sheet = getSheet(type);
    const rows = sheet.getDataRange().getValues();

    const validRows = rows.filter(row => row[0] instanceof Date && !isNaN(parseFloat(row[1])));

    for (const [dateMesure, poids] of validRows) {
      backup.appendRow([new Date(), 'MIGRATION', type, dateMesure, parseFloat(poids)]);
      count++;
    }
  }

  Logger.log(`✅ Migration terminée: ${count} entrée(s) copiée(s) vers la feuille "${CONFIG.BACKUP_SHEET_NAME}".`);
}
