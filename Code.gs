/**
 * Suivi de Poids - Application Google Apps Script
 * Stockage principal : Google Sheets (une feuille par profil)
 * Sauvegarde : feuille d'archive append-only
 */

const CONFIG = {
  TIMEZONE: "Europe/Paris",
  HISTORY_COUNT: 7,
  BACKUP_SHEET_NAME: "Sauvegarde",
  SHEETS: {
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
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Suivi de Poids')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = CONFIG.SHEETS[type];

  if (!config) {
    throw new Error("Type invalide : " + type);
  }

  for (const name of config.names) {
    const sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }

  return ss.insertSheet(config.names[0]);
}

function getBackupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.BACKUP_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.BACKUP_SHEET_NAME);
    sheet.appendRow(["horodatage_backup", "action", "type", "date_mesure", "poids"]);
  }

  return sheet;
}

function saveBackupEntry(action, type, date, weight) {
  const backupSheet = getBackupSheet();
  backupSheet.appendRow([new Date(), action, type, date, weight]);
}

function getWeightLimits(type) {
  const config = CONFIG.SHEETS[type];
  return {
    min: config.weightMin,
    max: config.weightMax
  };
}

function validateWeight(weight, type) {
  const parsed = parseFloat(String(weight).replace(',', '.'));
  const limits = getWeightLimits(type);

  if (isNaN(parsed)) {
    throw new Error("Valeur invalide : le poids doit être un nombre.");
  }

  if (parsed < limits.min || parsed > limits.max) {
    throw new Error(`Poids hors limites (${limits.min}-${limits.max} kg).`);
  }

  const decimals = type === 'cat' ? 100 : 10;
  return Math.round(parsed * decimals) / decimals;
}

function addWeight(weight, type) {
  const validWeight = validateWeight(weight, type);
  const sheet = getSheet(type);
  const measuredAt = new Date();

  sheet.appendRow([measuredAt, validWeight]);
  saveBackupEntry('ADD', type, measuredAt, validWeight);

  return getDashboardData(type);
}

function getDashboardData(type) {
  const sheet = getSheet(type);
  const data = sheet.getDataRange().getValues();
  const limits = getWeightLimits(type);

  const validData = data.filter(row =>
    row[0] instanceof Date &&
    !isNaN(parseFloat(row[1])) &&
    parseFloat(row[1]) >= limits.min &&
    parseFloat(row[1]) <= limits.max
  );

  if (validData.length === 0) {
    return {
      average: "--",
      history: [],
      range: "Aucune donnée",
      count: 0,
      min: "--",
      max: "--",
      trend: null,
      type: type
    };
  }

  validData.sort((a, b) => new Date(a[0]) - new Date(b[0]));

  const weights = validData.map(row => parseFloat(row[1]));
  const total = weights.reduce((acc, w) => acc + w, 0);

  const decimals = type === 'cat' ? 2 : 1;
  const avg = (total / weights.length).toFixed(decimals);
  const min = Math.min(...weights).toFixed(decimals);
  const max = Math.max(...weights).toFixed(decimals);

  let trend = null;
  if (validData.length >= 3) {
    const recentWeights = weights.slice(-3);
    const recentAvg = recentWeights.reduce((a, b) => a + b, 0) / recentWeights.length;
    const diff = recentAvg - parseFloat(avg);
    const threshold = type === 'cat' ? 0.1 : 0.3;
    trend = Math.abs(diff) > threshold ? (diff > 0 ? "up" : "down") : "stable";
  }

  const firstDate = validData[0][0];
  const lastDate = validData[validData.length - 1][0];
  const rangeStr = formatDateRange(firstDate, lastDate);

  const history = validData.slice(-CONFIG.HISTORY_COUNT).reverse().map(row => ({
    date: formatDateFR(row[0]),
    weight: parseFloat(row[1]).toFixed(decimals),
    fullDate: Utilities.formatDate(row[0], CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm")
  }));

  return {
    average: avg,
    history: history,
    range: rangeStr,
    count: validData.length,
    min: min,
    max: max,
    trend: trend,
    type: type
  };
}

function formatDateRange(firstDate, lastDate) {
  const first = formatMonthYearFR(firstDate);
  const last = formatMonthYearFR(lastDate);
  return first === last ? first : `${first} → ${last}`;
}

function formatDateFR(date) {
  const jours = ['Dim', 'Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam'];
  const mois = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc'];

  const jour = jours[date.getDay()];
  const numJour = String(date.getDate()).padStart(2, '0');
  const nomMois = mois[date.getMonth()];

  return `${jour} ${numJour} ${nomMois}`;
}

function formatMonthYearFR(date) {
  const mois = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc'];
  return `${mois[date.getMonth()]} ${date.getFullYear()}`;
}

function deleteLastEntry(type) {
  const sheet = getSheet(type);
  const lastRow = sheet.getLastRow();

  if (lastRow < 1) {
    throw new Error("Aucun relevé à supprimer.");
  }

  const [deletedDate, deletedWeight] = sheet.getRange(lastRow, 1, 1, 2).getValues()[0];
  sheet.deleteRow(lastRow);
  saveBackupEntry('DELETE', type, deletedDate, deletedWeight);

  return getDashboardData(type);
}
