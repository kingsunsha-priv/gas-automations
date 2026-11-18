// ---------------------------------------------------------
// 1. ENTRY POINTS
// ---------------------------------------------------------

/**
 * TEST MODE: Runs the scraper for only the very first category (Senior HS).
 * Outputs to a sheet named "Ranking Test".
 */
function testScrapeFirstCategory() {
  executeScraper(true);
}

/**
 * FULL MODE: Runs the scraper for all 30 categories.
 * Outputs to a sheet named "Ranking".
 */
function runFullScrape() {
  executeScraper(false);
}

// ---------------------------------------------------------
// 2. CORE SCRAPING LOGIC
// ---------------------------------------------------------

function executeScraper(isTestMode) {
  // Always fetch the latest cookie from the sheet and refresh the project's cache
  const cfg = App.forModule('rankings');
  const myCookies = cfg.get('web_cookie', true);
  const ogId = cfg.get('org_id', true);
  
  if (!myCookies) {
    Logger.log("Error: Cookie not found. Please ensure 'badminton_sweden_web_cookie' is configured in your settings sheet.");
    return;
  }

  // 2. HTTP Request Setup
  const options = {
    'method': 'get',
    'muteHttpExceptions': true,
    'headers': {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      'Cookie': myCookies,
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
      'Accept-Language': 'en-US,en;q=0.5'
    }
  };

  // 3. Automatically get the latest ranking ID
  Logger.log("Fetching the latest ranking ID from the overview page...");
  const rankingId = getLatestRankingId(options);
  
  if (!rankingId) {
    Logger.log("Failed to find a valid ranking ID. Script aborted.");
    return;
  }
  Logger.log(`Success! Found latest ranking ID: ${rankingId}`);
  
  Utilities.sleep(1500);

  // 4. Generate the full list of Categories
  const categoryGroups = [
    { name: "Senior", startId: 1091 },
    { name: "U13", startId: 1137 },
    { name: "U15", startId: 1142 },
    { name: "U17", startId: 1147 },
    { name: "U19", startId: 1152 },
    { name: "U22", startId: 1157 }
  ];
  const events = ["HS", "DS", "HD", "DD", "MD"];
  let targetCategories = [];

  categoryGroups.forEach(group => {
    events.forEach((event, index) => {
      targetCategories.push({
        groupName: group.name,
        eventName: event,
        categoryId: group.startId + index
      });
    });
  });

  if (isTestMode) {
    targetCategories = targetCategories.slice(0, 1);
    Logger.log("--- TEST MODE ENABLED: Only scraping the first category. ---");
  }

  let allParsedData = [];
  
  // Define our clean, hardcoded headers exactly how we want them
  const headers = [
    "Age Group", "Event", "Placering", "Rank Change", "Spelare", 
    "Licnr", "Födelseår", "Poäng", "Tävlingar", "Distrikt", "Klubb", "Nivå"
  ];
  allParsedData.push(headers);

  // 5. Loop through categories
  Logger.log(`Starting scrape of ${targetCategories.length} categories using ID ${rankingId}...`);
  
  for (let i = 0; i < targetCategories.length; i++) {
    const cat = targetCategories[i];
    const url = `https://badmintonsweden.tournamentsoftware.com/ranking/category.aspx?id=${rankingId}&category=${cat.categoryId}&ogid=${ogId}&p=1&ps=100`;
    
    Logger.log(`[${i+1}/${targetCategories.length}] Fetching ${cat.groupName} ${cat.eventName}...`);
    
    const response = UrlFetchApp.fetch(url, options);
    const htmlCode = response.getContentText();

    if (response.getResponseCode() !== 200 || htmlCode.includes("Cloudflare")) {
      Logger.log(`Failed to fetch ${cat.groupName} ${cat.eventName}. Cookies might be expired.`);
      break; 
    }

    const $ = Cheerio.load(htmlCode);
    
    $('table.ruler tr').each((rowIndex, rowElement) => {
      const isHeaderRow = $(rowElement).find('th').length > 0;
      if (isHeaderRow) return; // Skip HTML headers, we hardcoded our own
      
      const combinedRowText = $(rowElement).text().toLowerCase();
      // Filter out Pagination rows
      if (combinedRowText.includes("resultat per sida") || 
          combinedRowText.includes("results per page") ||
          combinedRowText.includes("nästa") ||
          combinedRowText.includes("sista")) {
        return; 
      }

      const tds = $(rowElement).find('td');
      
      // Ensure this is a valid player data row (should have ~12 standard cells)
      if (tds.length >= 11) {
        
        // Extract Rank Change with simplified format (x, -x, 0, New)
        const rankChangeEl = $(tds[1]);
        let rankChange = rankChangeEl.text().replace(/\u00a0/g, " ").trim();
        const className = rankChangeEl.attr('class') || "";
        
        if (className.includes('rank_down')) rankChange = `-${rankChange}`;
        else if (className.includes('rank_up')) rankChange = rankChange; // Keep as positive number
        else if (className.includes('rank_same')) rankChange = "0";
        else if (className.includes('rank_new')) rankChange = "New";

        // Extract explicit cell indexes (ignoring indices 2 and 4 which hold empty icons/flags)
        const placering = $(tds[0]).text().replace(/\u00a0/g, " ").trim();
        const spelare   = $(tds[3]).text().replace(/\u00a0/g, " ").trim();
        const licnr     = $(tds[5]).text().replace(/\u00a0/g, " ").trim();
        const fodelsear = $(tds[6]).text().replace(/\u00a0/g, " ").trim();
        const poang     = $(tds[7]).text().replace(/\u00a0/g, " ").trim();
        const tavlingar = $(tds[8]).text().replace(/\u00a0/g, " ").trim();
        const distrikt  = $(tds[9]).text().replace(/\u00a0/g, " ").trim();
        const klubb     = $(tds[10]).text().replace(/\u00a0/g, " ").trim();
        const niva      = tds.length >= 12 ? $(tds[11]).text().replace(/\u00a0/g, " ").trim() : "";

        // Push cleanly compiled row
        allParsedData.push([
          cat.groupName, // Age Group
          cat.eventName, // Event
          placering,
          rankChange,
          spelare,
          licnr,
          fodelsear,
          poang,
          tavlingar,
          distrikt,
          klubb,
          niva
        ]);
      }
    });

    Utilities.sleep(1500); 
  }

  // 6. Post-process and write data
  if (allParsedData.length > 1) { 
    const targetSheetName = isTestMode ? "Ranking Test" : "Ranking";
    writeToSheet(allParsedData, targetSheetName);
  } else {
    Logger.log("No player data was found. The tables might be empty or layout changed.");
  }
}

// ---------------------------------------------------------
// 3. HELPER FUNCTIONS
// ---------------------------------------------------------

/**
 * Scrape the latest ranking ID from the overview page.
 */
function getLatestRankingId(options) {
  const url = "https://badmintonsweden.tournamentsoftware.com/ranking";
  const response = UrlFetchApp.fetch(url, options);
  const htmlCode = response.getContentText();

  if (response.getResponseCode() !== 200 || htmlCode.includes("Cloudflare")) {
    Logger.log("Blocked by Cloudflare while trying to fetch the ranking overview.");
    return null;
  }

  const $ = Cheerio.load(htmlCode);
  let latestId = null;

  $('a').each((i, el) => {
    const href = $(el).attr('href');
    if (href && href.match(/[?&]id=(\d+)/i)) {
      const match = href.match(/[?&]id=(\d+)/i);
      if (match && match[1]) {
        latestId = match[1]; 
        return false; 
      }
    }
  });

  return latestId;
}

/**
 * Dedicated function to handle writing data to a Google Sheet.
 */
function writeToSheet(data, sheetName) {
  Logger.log(`Preparing to write data to sheet: "${sheetName}"`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found. Creating a new one.`);
    sheet = ss.insertSheet(sheetName);
  }

  const maxCols = Math.max(...data.map(row => row.length));

  sheet.clear(); 
  sheet.getRange(1, 1, data.length, maxCols).setValues(data);
  
  Logger.log(`Success! Wrote ${data.length - 1} players to the "${sheetName}" sheet.`);
}
