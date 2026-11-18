// =========================================================
// WEB APP ENDPOINTS FOR RANKING & SJT PLAYERS
// =========================================================

const RANKING_CACHE_KEY = 'lugi_public_rankings_v1';
const RANKING_CACHE_TTL = 21600; // 6 hours

/**
 * Serves the public ranking page with pre-injected JSON data
 * (avoids an extra google.script.run round-trip on load).
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  const initialPayload = getPublicRankingData(false);
  template.initialDataJson = JSON.stringify(initialPayload).replace(/</g, '\\u003c');

  return template
    .evaluate()
    .setTitle('Lugi Badminton - Ranking & SJT-spelare')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Reads "SJT Players" and "Ranking" sheets (with CacheService fast path).
 * @param {boolean=} forceRefresh Bypass cache and reload from sheets.
 */
function getPublicRankingData(forceRefresh) {
  const cache = CacheService.getScriptCache();

  if (!forceRefresh) {
    const cached = cache.get(RANKING_CACHE_KEY);
    if (cached) {
      return JSON.parse(cached);
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sjtSheet = ss.getSheetByName('SJT Players');
  const rankingSheet = ss.getSheetByName('Ranking');

  const sjtData = readSheetTable_(sjtSheet);
  const rankingData = readSheetTable_(rankingSheet);

  const payload = {
    sjt: sjtData,
    ranking: rankingData,
    lastUpdated: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() || 'Europe/Stockholm',
      'yyyy-MM-dd HH:mm'
    ),
  };

  try {
    cache.put(RANKING_CACHE_KEY, JSON.stringify(payload), RANKING_CACHE_TTL);
  } catch (e) {
    // Ignore if payload exceeds 100 KB CacheService limit
  }

  return payload;
}

/**
 * Helper to read a sheet as { headers: string[], rows: string[][] }.
 */
function readSheetTable_(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) {
    return { headers: [], rows: [] };
  }
  const values = sheet.getDataRange().getDisplayValues();
  return {
    headers: values[0],
    rows: values.slice(1).filter((r) => r.some((cell) => String(cell).trim() !== '')),
  };
}

/**
 * Fast action: Re-runs extractSJTPlayers() from the existing "Ranking" sheet
 * and refreshes the Web App cache (~1-2 seconds).
 */
function rebuildSjtAndGetRankings() {
  if (typeof extractSJTPlayers === 'function') {
    extractSJTPlayers();
    SpreadsheetApp.flush();
  }
  return getPublicRankingData(true);
}

/**
 * Full sync action: Scrapes all 30 categories via runFullScrape(),
 * updates "SJT Players" via extractSJTPlayers(), and refreshes the Web App cache (~45-60s).
 */
function syncAllRankingsAndGet() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('En rankinguppdatering pågår redan. Vänta en stund och försök igen.');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const beforeCount = ss.getSheetByName('Ranking')
      ? ss.getSheetByName('Ranking').getLastRow()
      : 0;

    runFullScrape();
    SpreadsheetApp.flush();

    const afterSheet = ss.getSheetByName('Ranking');
    if (!afterSheet || afterSheet.getLastRow() <= 1) {
      throw new Error(
        'Kunde inte hämta rankingdata (TournamentSoftware-cookien kan ha gått ut). Uppdatera cookien i inställningsarket.'
      );
    }

    if (typeof extractSJTPlayers === 'function') {
      extractSJTPlayers();
      SpreadsheetApp.flush();
    }

    return getPublicRankingData(true);
  } finally {
    lock.releaseLock();
  }
}