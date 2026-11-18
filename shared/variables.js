const App = (function () {
  // 1. TRUE PRIVATE CLOSURE STATE
  let _spreadsheet = null;
  let _sheets = {};
  let _settingsMap = null;
  let _badmintonCalendar = null;
  let _htmlns = null;

  // Master configuration spreadsheet ID
  const CONFIG_SPREADSHEET_ID = '1epFzmeYQEUjck-ty6GVpjJm33PTrIf7bY1mEBrcREbI';
  const CONFIG_SHEET_NAME = 'Settings';
  const CACHE_KEY = 'central_app_settings_v1';
  const CACHE_TTL = 21600; // 6 hours

  function loadSettings(bypassCache) {
    if (_settingsMap && !bypassCache) {
      return _settingsMap;
    }

    const cache = CacheService.getScriptCache();
    if (!bypassCache) {
      const cachedJson = cache.get(CACHE_KEY);
      if (cachedJson) {
        _settingsMap = JSON.parse(cachedJson);
        return _settingsMap;
      }
    }

    const ss = CONFIG_SPREADSHEET_ID
      ? SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);

    _settingsMap = {};
    if (sheet) {
      const rows = sheet.getDataRange().getValues();
      // Expected columns: [0] Module, [1] Key, [2] Value
      for (let i = 1; i < rows.length; i++) {
        const mod = String(rows[i][0] || 'global').trim().toLowerCase();
        const key = String(rows[i][1] || '').trim();
        if (key) {
          _settingsMap[`${mod}.${key}`] = rows[i][2];
        }
      }

      try {
        cache.put(CACHE_KEY, JSON.stringify(_settingsMap), CACHE_TTL);
      } catch (e) {
        // Ignore if payload exceeds 100 KB
      }
    }

    return _settingsMap;
  }

  return {
    get spreadsheet() {
      if (!_spreadsheet) {
        _spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      }
      return _spreadsheet;
    },

    getSheet(sheetName) {
      if (!_sheets[sheetName]) {
        _sheets[sheetName] = this.spreadsheet.getSheetByName(sheetName);
      }
      return _sheets[sheetName];
    },

    /**
     * Direct lookup by full namespaced key (e.g. "rankings.web_cookie" or "global.calendar_id").
     */
    getSetting(fullKey, bypassCache = false) {
      const settings = loadSettings(bypassCache);
      return settings[fullKey];
    },

    /**
     * Returns a module-scoped config reader that checks "<module>.<key>"
     * first and automatically falls back to "global.<key>".
     */
    forModule(moduleName) {
      const prefix = String(moduleName || 'global').trim().toLowerCase();
      return {
        get(key, bypassCache = false) {
          const settings = loadSettings(bypassCache);
          const namespacedKey = `${prefix}.${key}`;
          return settings[namespacedKey] !== undefined
            ? settings[namespacedKey]
            : settings[`global.${key}`];
        },
      };
    },

    clearSettingsCache() {
      _settingsMap = null;
      CacheService.getScriptCache().remove(CACHE_KEY);
    },

    // Shared lazy services
    get badmintonCalendar() {
      if (!_badmintonCalendar) {
        _badmintonCalendar = CalendarApp.getCalendarById(this.getSetting('global.calendar_id'));
      }
      return _badmintonCalendar;
    },

    get htmlns() {
      if (!_htmlns) {
        _htmlns = XmlService.getNamespace('http://www.w3.org/1999/xhtml');
      }
      return _htmlns;
    },
  };
})();