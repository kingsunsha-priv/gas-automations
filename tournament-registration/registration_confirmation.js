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
  INCLUDE_PAST_TOURNAMENTS: false, // Set to true if you also want ended tournaments (e.g., 2026-11-07)
  MAX_THREADS: 50,
  SHEET_NAME: 'Tournament Entries',

  // Lazy-loaded spreadsheet reference
  _spreadsheet: null,
  get spreadsheet() {
    if (!this._spreadsheet) {
      this._spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }
    return this._spreadsheet;
  },

  /**
   * Orchestrates Gmail fetching, deduplication, HTML parsing, and Sheet updates.
   */
  run() {
    const latestMessages = this.fetchLatestMessagesByTournament();

    if (latestMessages.size === 0) {
      Logger.log('No matching Badminton Sweden confirmation emails found.');
      return;
    }

    const todayStart = new Date();
    todayStart.setHours(0, 0, 0, 0);

    const parsedTournaments = [];
    latestMessages.forEach((entry) => {
      const parsed = this.parseConfirmationEmail(entry.message, entry.tournamentId);

      if (!this.INCLUDE_PAST_TOURNAMENTS && parsed.endDate && parsed.endDate < todayStart) {
        Logger.log(`Skipping past tournament: ${parsed.tournamentName} (${parsed.tournamentDate})`);
        return;
      }

      const activeCount = parsed.entries.filter((e) => !e.withdrawn).length;
      const withdrawnCount = parsed.entries.filter((e) => e.withdrawn).length;
      Logger.log(
        `Parsed "${parsed.tournamentName}" (${parsed.tournamentDate}) -> Confirmed: ${activeCount}, Withdrawn: ${withdrawnCount}`
      );

      parsedTournaments.push(parsed);
    });

    if (parsedTournaments.length === 0) {
      Logger.log('No current or upcoming tournaments found to sync.');
      return;
    }

    this.syncEntriesSheet(parsedTournaments);
  },

  /**
   * Searches recent Gmail confirmations and keeps ONLY the newest message per Tournament ID.
   */
  fetchLatestMessagesByTournament() {
    const query = `from:notifications@tournamentsoftware.com subject:"bekräftelse av anmälan till" newer_than:${this.LOOKBACK_DAYS}d`;
    const threads = GmailApp.search(query, 0, this.MAX_THREADS);
    const latestByTournament = new Map();

    threads.forEach((thread) => {
      const messages = thread.getMessages();
      messages.forEach((msg) => {
        const html = msg.getBody();
        const tournamentId = this.extractTournamentId(html, msg.getSubject());
        const msgDate = msg.getDate();

        const existing = latestByTournament.get(tournamentId);
        if (!existing || msgDate.getTime() > existing.date.getTime()) {
          latestByTournament.set(tournamentId, {
            tournamentId,
            date: msgDate,
            message: msg,
          });
        }
      });
    });

    return latestByTournament;
  },

  /**
   * Extracts the 36-character TournamentSoftware GUID from URL-decoded links in the email.
   */
  extractTournamentId(html, subjectFallback) {
    const $ = Cheerio.load(html);
    let foundGuid = null;

    $('a[href]').each((i, el) => {
      let href = $(el).attr('href') || '';
      try {
        href = decodeURIComponent(href);
      } catch (e) {
        // Keep raw href if URI decoding fails
      }

      if (href.toLowerCase().includes('tournamentsoftware.com')) {
        const guidMatch = href.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);
        if (guidMatch) {
          foundGuid = guidMatch[0].toUpperCase();
          return false;
        }
      }
    });

    if (foundGuid) return foundGuid;

    const cleanName = $('a.mail-media__link').first().text().replace(/\u00a0/g, ' ').trim();
    return cleanName || subjectFallback.replace(/^.*bekräftelse av anmälan till\s*/i, '').trim();
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
   * Syncs all confirmed and withdrawn entries into one sheet, overwriting
   * existing rows for each refreshed tournament ID.
   */
  syncEntriesSheet(parsedTournaments) {
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

    const retainedRows = existingData
      .slice(1)
      .filter((row) => row[0] && !tournamentsToOverwrite.has(String(row[0]).trim()));

    const finalRows = retainedRows.concat(newRows);

    // Sort by Tournament Date, Tournament Name, Active (false) before Withdrawn (true), then Class
    finalRows.sort((a, b) => {
      return (
        String(a[2]).localeCompare(String(b[2])) ||
        String(a[1]).localeCompare(String(b[1])) ||
        Number(Boolean(a[3])) - Number(Boolean(b[3])) ||
        String(a[4]).localeCompare(String(b[4]))
      );
    });

    this.writeBatchToSheet(sheet, headers, finalRows);
    Logger.log(
      `Synced ${this.SHEET_NAME}: updated ${tournamentsToOverwrite.size} tournament(s), total ${finalRows.length} rows.`
    );
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