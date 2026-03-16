/**
 * Script de migration Google Sheets → Firebase Firestore
 * -------------------------------------------------------
 * AVANT D'EXÉCUTER :
 *   1. Dans Firebase Console → Firestore → Rules, mets temporairement :
 *        allow read, write: if true;
 *   2. Ouvre ce script dans Google Apps Script (extensions.google.com/appscript)
 *   3. Exécute la fonction  migrerVersFirebase()
 *   4. Une fois la migration terminée, restreins tes règles Firestore
 */

const FIREBASE_PROJECT_ID = "poids-79bd4";
const FIREBASE_API_KEY    = "AIzaSyDMiLaxqORBjdm-Tkq56F8cEKbxVW3H8Pc";

const SHEETS_CONFIG = {
  human: {
    names: ["Poids Anthony", "Anthony", "Feuille 1", "Feuille1", "Sheet1"],
    weightMin: 30,
    weightMax: 250
  },
  cat: {
    names: ["Poids Minou", "Minou", "Chat"],
    weightMin: 1,
    weightMax: 15
  }
};

/**
 * Fonction principale — exécute-la une seule fois
 */
function migrerVersFirebase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalMigré = 0;
  let erreurs = 0;

  for (const type of ['human', 'cat']) {
    const config = SHEETS_CONFIG[type];
    let sheet = null;

    for (const name of config.names) {
      sheet = ss.getSheetByName(name);
      if (sheet) break;
    }

    if (!sheet) {
      Logger.log(`⚠️  Feuille non trouvée pour le type "${type}" — ignoré`);
      continue;
    }

    const data = sheet.getDataRange().getValues();
    const valides = data.filter(row =>
      row[0] instanceof Date &&
      !isNaN(parseFloat(row[1])) &&
      parseFloat(row[1]) >= config.weightMin &&
      parseFloat(row[1]) <= config.weightMax
    );

    Logger.log(`📊  ${type} → feuille "${sheet.getName()}" : ${valides.length} entrée(s) valide(s)`);

    for (const row of valides) {
      const date   = row[0];
      const weight = parseFloat(row[1]);

      try {
        creerDocument(type, weight, date.toISOString());
        totalMigré++;
        Utilities.sleep(120); // pause pour éviter le rate-limiting Firebase
      } catch (e) {
        erreurs++;
        Logger.log(`❌  Erreur (${type}, ${weight} kg, ${date}) : ${e.message}`);
      }
    }

    Logger.log(`✅  ${type} : migration terminée`);
  }

  Logger.log(`\n🎉 Terminé — ${totalMigré} entrée(s) migrée(s), ${erreurs} erreur(s)`);
}

/**
 * Crée un document dans Firestore via l'API REST
 */
function creerDocument(collection, weight, isoDate) {
  const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/${collection}?key=${FIREBASE_API_KEY}`;

  const body = {
    fields: {
      weight:    { doubleValue: weight },
      timestamp: { timestampValue: isoDate }
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error(`HTTP ${code} — ${response.getContentText()}`);
  }
}

/**
 * Vérifie que la connexion à Firestore fonctionne (test rapide)
 */
function testerConnexion() {
  try {
    creerDocument('_test', 0, new Date().toISOString());
    Logger.log('✅ Connexion Firestore OK');
  } catch (e) {
    Logger.log('❌ Connexion échouée : ' + e.message);
  }
}
