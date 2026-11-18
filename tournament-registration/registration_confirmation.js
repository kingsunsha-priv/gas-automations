// =========================================================
// 1. GLOBAL ENTRY POINTS (For Menus & Time-Driven Triggers)
// =========================================================

/**
 * Main entry point: Fetches the latest Badminton Sweden entry confirmation
 * email per tournament from Gmail and syncs both confirmed and withdrawn
 * entries into a single Google Sheet.
 */
function syncTournamentEntriesFromGmail() {
  TournamentEntrySync.run();
}

// =========================================================
// 2. NAMESPACED APPLICATION LOGIC
// =========================================================

const TournamentEntrySync = {
  // Configuration
  LOOKBACK_DAYS: 60,               // Search Gmail messages from the last N days
  INCLUDE_PAST_TOURNAMENTS: false, // Set to true if you also want ended tournaments
  MAX_THREADS: 50,
  SHEET_NAME: 'Tournament Entries',
  MSG_META_CACHE_KEY: 'ts_msg_meta_v2',

  // Lazy-loaded spreadsheet reference
  _spreadsheet: null,
  get spreadsheet() {
    if (!this._spreadsheet) {
      this._spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }
    return this._spreadsheet;
  },

  isGuid(val) {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
      String(val || '').trim()
    );
  },

  /**
   * Orchestrates Gmail fetching, incremental change detection, HTML parsing, and Sheet updates.
   */
  run() {
    const latestMessages = this.fetchLatestMessagesByTournament();

    if (latestMessages.size === 0) {
      Logger.log('No matching Badminton Sweden confirmation emails found.');
      return;
    }

    const headers = [
      'Tournament ID',
      'Tournament Name',
      'Tournament Date',
      'Withdrawn',
      'Class',
      'Player 1',
      'License 1',
      'Country 1',
      'Gender 1',
      'Player 2',
      'License 2',
      'Country 2',
      'Gender 2',
      'Request Status',
      'Removed On',
      'Withdrawal Reason',
      'Total Tournament Cost',
      'Email Timestamp',
    ];

    const sheet = this.getOrCreateSheet(this.SHEET_NAME, headers);
    const existingData = sheet.getDataRange().getValues();
    const existingTimestamps = this.getExistingTimestampMap(existingData, 0, headers.length - 1);

    // Force a clean rewrite if any legacy non-GUID rows (e.g., subject-fallback duplicates) exist
    const hasInvalidIdRows = existingData
      .slice(1)
      .some((row) => row[0] && !this.isGuid(row[0]));

    const parsedTournaments = [];
    latestMessages.forEach((entry) => {
      const savedTime = existingTimestamps.get(entry.tournamentId) || 0;

      // Skip downloading/parsing full HTML body if already synced and no legacy cleanup is needed
      if (!hasInvalidIdRows && entry.date.getTime() <= savedTime + 1000) {
        return;
      }

      const parsed = this.parseConfirmationEmail(entry.message, entry.tournamentId);

      const activeCount = parsed.entries.filter((e) => !e.withdrawn).length;
      const withdrawnCount = parsed.entries.filter((e) => e.withdrawn).length;
      Logger.log(
        `Parsed "${parsed.tournamentName}" (${parsed.tournamentDate}) -> Confirmed: ${activeCount}, Withdrawn: ${withdrawnCount}`
      );

      parsedTournaments.push(parsed);
    });

    if (parsedTournaments.length === 0) {
      Logger.log('Tournament Entries sheet is already up to date.');
      return;
    }

    this.syncEntriesSheet(sheet, headers, existingData, parsedTournaments);
  },

  /**
   * Searches recent Gmail confirmations and keeps ONLY the newest message per Tournament ID,
   * caching message metadata in CacheService so already-seen emails don't re-download bodies.
   */
  fetchLatestMessagesByTournament() {
    const query = `from:notifications@tournamentsoftware.com subject:"bekräftelse av anmälan till" newer_than:${this.LOOKBACK_DAYS}d`;
    const threads = GmailApp.search(query, 0, this.MAX_THREADS);
    const latestByTournament = new Map();

    const cache = CacheService.getScriptCache();
    let msgMeta = {};
    try {
      const rawMeta = cache.get(this.MSG_META_CACHE_KEY);
      if (rawMeta) msgMeta = JSON.parse(rawMeta);
    } catch (e) {
      msgMeta = {};
    }
    let metaChanged = false;

    const todayStart = new Date();
    todayStart.setHours(0, 0, 0, 0);
    const todayMs = todayStart.getTime();

    threads.forEach((thread) => {
      const messages = thread.getMessages();
      if (messages.length === 0) return;

      // Gmail orders messages in a thread chronologically; only inspect the newest message
      const msg = messages[messages.length - 1];
      const msgId = msg.getId();
      const msgDate = msg.getDate();

      let info = msgMeta[msgId];
      if (!info || !this.isGuid(info.tournamentId)) {
        // Strip inline base64 logos so the full HTML structure is lightweight and intact
        const cleanHtml = msg.getBody().replace(/data:image\/[^;]+;base64,[A-Za-z0-9+/=]+/gi, '');
        const tournamentId = this.extractTournamentId(cleanHtml, msg.getSubject());

        const $ = Cheerio.load(cleanHtml);
        const tournamentDate = ($('span.text-label--sub').first().text() || '')
          .replace(/\u00a0/g, ' ')
          .trim();
        const endDate = this.parseEndDate(tournamentDate);

        info = {
          tournamentId,
          endDateMs: endDate ? endDate.getTime() : null,
        };
        msgMeta[msgId] = info;
        metaChanged = true;
      }

      if (!this.INCLUDE_PAST_TOURNAMENTS && info.endDateMs && info.endDateMs < todayMs) {
        return;
      }

      const existing = latestByTournament.get(info.tournamentId);
      if (!existing || msgDate.getTime() > existing.date.getTime()) {
        latestByTournament.set(info.tournamentId, {
          tournamentId: info.tournamentId,
          date: msgDate,
          message: msg,
        });
      }
    });

    if (metaChanged) {
      try {
        cache.put(this.MSG_META_CACHE_KEY, JSON.stringify(msgMeta), 21600);
      } catch (e) {
        // Ignore cache write errors if payload exceeds 100 KB
      }
    }

    return latestByTournament;
  },

  /**
   * Extracts the 36-character TournamentSoftware GUID from URL-decoded links across the full email.
   */
  extractTournamentId(cleanHtml, subjectFallback) {
    const hrefRegex = /href\s*=\s*["']([^"']+)["']/gi;
    let match;

    while ((match = hrefRegex.exec(cleanHtml)) !== null) {
      let href = match[1];
      try {
        href = decodeURIComponent(href);
      } catch (e) {
        // Keep raw href if URI decoding fails
      }

      if (href.toLowerCase().includes('tournamentsoftware.com')) {
        const guidMatch = href.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);
        if (guidMatch) {
          return guidMatch[0].toUpperCase();
        }
      }
    }

    const $ = Cheerio.load(cleanHtml);
    const cleanName = $('a.mail-media__link')
      .first()
      .text()
      .replace(/\u00a0/g, ' ')
      .replace(/,+\s*$/, '')
      .trim();

    if (cleanName) return cleanName;

    return (subjectFallback || '')
      .replace(/^.*bekräftelse av anmälan till\s*/i, '')
      .replace(/,+\s*$/, '')
      .trim();
  },

  /**
   * Parses both confirmed and withdrawn ("Inaktivera") entries from a confirmation email,
   * regardless of whether table headers use <th> or <td> tags.
   */
  parseConfirmationEmail(message, tournamentId) {
    const html = message.getBody();
    const emailDate = message.getDate();
    const $ = Cheerio.load(html);

    const clean = (str) => (str || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();

    const splitCellLines = (tdEl) => {
      const rawHtml = $(tdEl).html() || '';
      return rawHtml
        .split(/<br\s*\/?>/i)
        .map((fragment) => clean(Cheerio.load(fragment).text()))
        .filter((text) => text.length > 0);
    };

    const tournamentName = clean($('a.mail-media__link').first().text());
    const tournamentDate = clean($('span.text-label--sub').first().text());
    const endDate = this.parseEndDate(tournamentDate);

    // Inspect all leaf tables (tables that do not wrap outer layout tables)
    const leafTables = $('table').filter((i, el) => $(el).find('table').length === 0);

    const entries = [];
    const seenEntryKeys = new Set();
    let totalCost = '';

    leafTables.each((tIdx, tableEl) => {
      $(tableEl)
        .find('tr')
        .each((rIdx, tr) => {
          // Skip explicit <th> header rows
          if ($(tr).children('th').length > 0) return;

          const tds = $(tr).children('td');

          // Capture total cost footer row
          if (tds.length === 1) {
            const rowText = clean($(tds[0]).text());
            if (rowText.startsWith('Total entry costs:')) {
              totalCost = rowText.replace('Total entry costs:', '').trim();
            }
            return;
          }

          if (tds.length < 6) return;

          const klass = clean($(tds[0]).text());
          const klassLower = klass.toLowerCase();

          // Skip <td>-based header rows and email boilerplate
          if (!klass || klassLower === 'klass' || klassLower === 'class' || klass.length > 60) {
            return;
          }

          const licenses = splitCellLines(tds[1]);
          const countries = splitCellLines(tds[2]);
          const players = splitCellLines(tds[3]);
          const genders = splitCellLines(tds[4]);
          const col5Text = clean($(tds[5]).text());

          if (!players[0]) return;

          // Withdrawn ("Inaktivera") rows have 7+ columns and a YYYY-MM-DD timestamp in column 6 (index 5)
          const isWithdrawnRow = tds.length >= 7 && /\d{4}-\d{2}-\d{2}/.test(col5Text);

          const entryObj = isWithdrawnRow
            ? {
                withdrawn: true,
                klass,
                player1: players[0] || '',
                license1: licenses[0] || '',
                country1: countries[0] || '',
                gender1: genders[0] || '',
                player2: players[1] || '',
                license2: licenses[1] || '',
                country2: countries[1] || '',
                gender2: genders[1] || '',
                requestStatus: '',
                removedOn: col5Text,
                reason: clean($(tds[6]).text()),
              }
            : {
                withdrawn: false,
                klass,
                player1: players[0] || '',
                license1: licenses[0] || '',
                country1: countries[0] || '',
                gender1: genders[0] || '',
                player2: players[1] || '',
                license2: licenses[1] || '',
                country2: countries[1] || '',
                gender2: genders[1] || '',
                requestStatus: col5Text,
                removedOn: '',
                reason: '',
              };

          // Prevent duplicates if the email HTML includes a hidden mobile version of the same table
          const dedupeKey = [
            entryObj.withdrawn,
            entryObj.klass,
            entryObj.license1,
            entryObj.player1,
            entryObj.license2,
            entryObj.player2,
            entryObj.removedOn,
          ].join('|');

          if (!seenEntryKeys.has(dedupeKey)) {
            seenEntryKeys.add(dedupeKey);
            entries.push(entryObj);
          }
        });
    });

    return {
      tournamentId,
      tournamentName,
      tournamentDate,
      endDate,
      emailDate,
      totalCost,
      entries,
    };
  },

  /**
   * Parses the end Date object from strings like "2026-11-07" or "2026-01-24 till 2026-01-25".
   */
  parseEndDate(dateStr) {
    const matches = (dateStr || '').match(/\d{4}-\d{2}-\d{2}/g);
    if (!matches || matches.length === 0) return null;

    const lastDateStr = matches[matches.length - 1];
    const parts = lastDateStr.split('-').map(Number);
    return new Date(parts[0], parts[1] - 1, parts[2]);
  },

  /**
   * Syncs confirmed and withdrawn entries into one sheet, overwriting
   * existing rows only for tournaments that received a newer email.
   */
  syncEntriesSheet(sheet, headers, existingData, parsedTournaments) {
    const tournamentsToOverwrite = new Set();
    const newRows = [];

    parsedTournaments.forEach((t) => {
      tournamentsToOverwrite.add(t.tournamentId);

      t.entries.forEach((e) => {
        newRows.push([
          t.tournamentId,
          t.tournamentName,
          t.tournamentDate,
          e.withdrawn,
          e.klass,
          e.player1,
          e.license1,
          e.country1,
          e.gender1,
          e.player2,
          e.license2,
          e.country2,
          e.gender2,
          e.requestStatus,
          e.removedOn,
          e.reason,
          t.totalCost,
          t.emailDate,
        ]);
      });
    });

    // Drop rows being overwritten AND any legacy non-GUID fallback rows
    const retainedRows = existingData.slice(1).filter((row) => {
      const rowId = String(row[0] || '').trim();
      return rowId && this.isGuid(rowId) && !tournamentsToOverwrite.has(rowId);
    });

    const finalRows = retainedRows.concat(newRows);

    // Sort by Tournament Date, Tournament Name, Class, then Player 1
    finalRows.sort((a, b) => {
      return (
        String(a[2]).localeCompare(String(b[2])) ||
        String(a[1]).localeCompare(String(b[1])) ||
        String(a[4]).localeCompare(String(b[4])) ||
        String(a[5]).localeCompare(String(b[5]))
      );
    });

    this.writeBatchToSheet(sheet, headers, finalRows);
    CacheService.getScriptCache().remove(PUBLIC_ENTRIES_CACHE_KEY);
    Logger.log(
      `Synced ${this.SHEET_NAME}: updated ${tournamentsToOverwrite.size} tournament(s), total ${finalRows.length} rows.`
    );
  },

  getExistingTimestampMap(sheetValues, idColIndex, timestampColIndex) {
    const map = new Map();
    for (let i = 1; i < sheetValues.length; i++) {
      const id = String(sheetValues[i][idColIndex] || '').trim();
      const rawDate = sheetValues[i][timestampColIndex];
      if (!id || !rawDate) continue;

      const time = rawDate instanceof Date ? rawDate.getTime() : new Date(rawDate).getTime();
      if (!isNaN(time) && (!map.has(id) || time > map.get(id))) {
        map.set(id, time);
      }
    }
    return map;
  },

  getOrCreateSheet(sheetName, headers) {
    let sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = this.spreadsheet.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    return sheet;
  },

  writeBatchToSheet(sheet, headers, rows) {
    const lastRow = sheet.getLastRow();
    const clearCols = Math.max(sheet.getLastColumn(), headers.length);
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, clearCols).clearContent();
    }

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
  },
};

// =========================================================
// 3. WEB APP ENDPOINTS (Cached + Single-Trip Injection)
// =========================================================

const PUBLIC_ENTRIES_CACHE_KEY = 'public_tournament_entries_v1';
const PUBLIC_ENTRIES_CACHE_TTL = 21600; // 6 hours

/**
 * Serves the public registration page with data pre-injected into the HTML template,
 * eliminating the second google.script.run round-trip on page load.
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  const initialPayload = getPublicTournamentEntries(false);
  template.initialDataJson = JSON.stringify(initialPayload).replace(/</g, '\\u003c');

  return template
    .evaluate()
    .setTitle('Lugi Badminton - Tävlingsanmälningar')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Returns sanitized entry rows from CacheService (fast path) or "Tournament Entries" sheet.
 * @param {boolean=} forceRefresh Bypass cache and reload directly from the sheet.
 */
function getPublicTournamentEntries(forceRefresh) {
  const cache = CacheService.getScriptCache();

  if (!forceRefresh) {
    const cachedJson = cache.get(PUBLIC_ENTRIES_CACHE_KEY);
    if (cachedJson) {
      return JSON.parse(cachedJson);
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TournamentEntrySync.SHEET_NAME);
  const lastRow = sheet ? sheet.getLastRow() : 0;

  if (!sheet || lastRow <= 1) {
    return { entries: [], lastUpdated: '' };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 18).getDisplayValues();
  const entries = [];
  let latestSync = '';

  // Column indices (0-based):
  // 1: Tournament Name, 2: Tournament Date, 3: Withdrawn, 4: Class,
  // 5: Player 1, 9: Player 2, 13: Request Status, 14: Removed On, 17: Email Timestamp
  for (let i = 0; i < values.length; i++) {
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

  const payload = {
    entries,
    lastUpdated: latestSync,
  };

  try {
    cache.put(PUBLIC_ENTRIES_CACHE_KEY, JSON.stringify(payload), PUBLIC_ENTRIES_CACHE_TTL);
  } catch (e) {
    // Ignore cache write errors if payload exceeds 100 KB
  }

  return payload;
}

/**
 * Triggered by the Web App "Synka & uppdatera" button.
 * Runs Gmail confirmation sync, rebuilds the cache, and returns fresh rows.
 */
function syncAndGetTournamentEntries() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('En synkronisering pågår redan. Vänta några sekunder och försök igen.');
  }

  try {
    TournamentEntrySync.run();
    SpreadsheetApp.flush();
    return getPublicTournamentEntries(true);
  } finally {
    lock.releaseLock();
  }
}