// =========================================================
// 3. WEB APP ENDPOINTS (For Google Sites Embed)
// =========================================================

/**
 * Serves the public registration status page for embedding in Google Sites.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Lugi Badminton - Tävlingsanmälningar')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Returns sanitized entry rows from "Tournament Entries" for the frontend UI.
 * Omits sensitive columns (License numbers, Tournament GUIDs, and Total Costs).
 */
function getPublicTournamentEntries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TournamentEntrySync.SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { entries: [], lastUpdated: '' };
  }

  // getDisplayValues() ensures timestamps and booleans serialize cleanly to HTML
  const values = sheet.getDataRange().getDisplayValues();
  const entries = [];
  let latestSync = '';

  // Column map from TournamentEntrySync.syncEntriesSheet:
  // 1: Tournament Name, 2: Tournament Date, 3: Withdrawn, 4: Class,
  // 5: Player 1, 9: Player 2, 13: Request Status, 14: Removed On, 17: Email Timestamp
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[1]) continue;

    const emailTs = row[17] || '';
    if (emailTs > latestSync) {
      latestSync = emailTs;
    }

    entries.push({
      tournamentName: row[1],
      tournamentDate: row[2],
      withdrawn: String(row[3]).toUpperCase() === 'TRUE',
      klass: row[4],
      player1: row[5],
      player2: row[9],
      status: row[13],
      removedOn: row[14],
    });
  }

  return {
    entries,
    lastUpdated: latestSync,
  };
}

/**
 * Triggered by the Web App "Synka & uppdatera" button.
 * Runs Gmail confirmation sync, flushes sheet writes, and returns updated rows.
 */
function syncAndGetTournamentEntries() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('En synkronisering pågår redan. Vänta några sekunder och försök igen.');
  }

  try {
    TournamentEntrySync.run();
    SpreadsheetApp.flush();
    return getPublicTournamentEntries();
  } finally {
    lock.releaseLock();
  }
}
