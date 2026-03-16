/**
 * Suivi de Poids - Application Google Apps Script
 * Version 7.0 - Deux onglets : Anthony & Minou
 */

const CONFIG = {
  TIMEZONE: "Europe/Paris",
  HISTORY_COUNT: 7,
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
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Suivi de Poids')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Récupère la feuille de données selon le type (human/cat)
 */
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
  
  // Créer la feuille si elle n'existe pas
  const newSheet = ss.insertSheet(config.names[0]);
  return newSheet;
}

/**
 * Récupère les limites de poids selon le type
 */
function getWeightLimits(type) {
  const config = CONFIG.SHEETS[type];
  return {
    min: config.weightMin,
    max: config.weightMax
  };
}

/**
 * Valide et nettoie une valeur de poids
 */
function validateWeight(weight, type) {
  const parsed = parseFloat(String(weight).replace(',', '.'));
  const limits = getWeightLimits(type);
  
  if (isNaN(parsed)) {
    throw new Error("Valeur invalide : le poids doit être un nombre.");
  }
  
  if (parsed < limits.min || parsed > limits.max) {
    throw new Error(`Poids hors limites (${limits.min}-${limits.max} kg).`);
  }
  
  // Arrondi à 1 décimale pour humain, 2 décimales pour chat
  const decimals = type === 'cat' ? 100 : 10;
  return Math.round(parsed * decimals) / decimals;
}

/**
 * Ajoute un relevé de poids
 */
function addWeight(weight, type) {
  const validWeight = validateWeight(weight, type);
  const sheet = getSheet(type);
  sheet.appendRow([new Date(), validWeight]);
  return getDashboardData(type);
}

/**
 * Récupère toutes les données pour le dashboard
 */
function getDashboardData(type) {
  const sheet = getSheet(type);
  const data = sheet.getDataRange().getValues();
  const limits = getWeightLimits(type);
  
  // Filtrer les données valides (date + nombre dans les limites)
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

  // Trier par date
  validData.sort((a, b) => new Date(a[0]) - new Date(b[0]));

  // Calculs statistiques
  const weights = validData.map(row => parseFloat(row[1]));
  const total = weights.reduce((acc, w) => acc + w, 0);
  
  // Précision selon le type
  const decimals = type === 'cat' ? 2 : 1;
  const avg = (total / weights.length).toFixed(decimals);
  const min = Math.min(...weights).toFixed(decimals);
  const max = Math.max(...weights).toFixed(decimals);

  // Tendance
  let trend = null;
  if (validData.length >= 3) {
    const recentWeights = weights.slice(-3);
    const recentAvg = recentWeights.reduce((a, b) => a + b, 0) / recentWeights.length;
    const diff = recentAvg - parseFloat(avg);
    const threshold = type === 'cat' ? 0.1 : 0.3;
    if (Math.abs(diff) > threshold) {
      trend = diff > 0 ? "up" : "down";
    } else {
      trend = "stable";
    }
  }

  // Plage de dates
  const firstDate = validData[0][0];
  const lastDate = validData[validData.length - 1][0];
  const rangeStr = formatDateRange(firstDate, lastDate);

  // Historique (dates en français)
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

/**
 * Formate la plage de dates
 */
function formatDateRange(firstDate, lastDate) {
  const first = formatMonthYearFR(firstDate);
  const last = formatMonthYearFR(lastDate);
  
  if (first === last) {
    return first;
  }
  return `${first} → ${last}`;
}

/**
 * Formate une date en français (ex: "Mar 03 Fév")
 */
function formatDateFR(date) {
  const jours = ['Dim', 'Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam'];
  const mois = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc'];
  
  const jour = jours[date.getDay()];
  const numJour = String(date.getDate()).padStart(2, '0');
  const nomMois = mois[date.getMonth()];
  
  return `${jour} ${numJour} ${nomMois}`;
}

/**
 * Formate mois et année en français (ex: "Fév 2025")
 */
function formatMonthYearFR(date) {
  const mois = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc'];
  return `${mois[date.getMonth()]} ${date.getFullYear()}`;
}

/**
 * Supprime le dernier relevé
 */
function deleteLastEntry(type) {
  const sheet = getSheet(type);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 1) {
    throw new Error("Aucun relevé à supprimer.");
  }
  
  sheet.deleteRow(lastRow);
  return getDashboardData(type);
}

/**
 * Récupère les données des deux onglets au chargement
 */
function getAllData() {
  return {
    human: getDashboardData('human'),
    cat: getDashboardData('cat')
  };
}
