// Global scope: Initialized once per execution session
var COUNTRY_MAP = {
  "SE": "瑞典", "SWEDEN": "瑞典", "瑞典": "瑞典",
  "DE": "德国", "GERMANY": "德国", "德国": "德国",
  "PL": "波兰", "POLAND": "波兰", "波兰": "波兰",
  "CZ": "捷克", "CZECHIA": "捷克", "CZECH REPUBLIC": "捷克", "捷克": "捷克",
  "DK": "丹麦", "DENMARK": "丹麦", "丹麦": "丹麦",
  "CN": "中国", "CHINA": "中国", "中国": "中国",
  "BY": "白俄罗斯", "BELARUS": "白俄罗斯", "白俄罗斯": "白俄罗斯",
  "RU": "俄罗斯", "RUSSIA": "俄罗斯", "俄罗斯": "俄罗斯",
  "KZ": "哈萨克斯坦", "KAZAKHSTAN": "哈萨克斯坦", "哈萨克斯坦": "哈萨克斯坦"
};

/**
 * Unifies various country inputs (ISO codes, English names, Chinese names) into a single standard format.
 */
function NormalizeCountry(country) {
  if (!country) return "";
  var key = String(country).trim().toUpperCase();
  return COUNTRY_MAP[key] || country;
}

/**
 * Creates a fast O(1) column lookup map from header array
 */
function buildHeaderMap(headers) {
  var map = {};
  for (var c = 0; c < headers.length; c++) {
    map[headers[c]] = c;
  }
  return map;
}

/**
 * Helper to reliably parse time, location, and status from logs
 */
function ParseLogLine(line) {
  if (!line) return null;
  if (line.indexOf('\t') >= 0) {
    var parts = line.split('\t');
    return {
      time: parts[0].trim(),
      location: parts.length > 2 ? parts[1].trim() : "", // 提取出 Location (如: Wolka Kosowska, Poland)
      status: parts[parts.length - 1].trim()
    };
  } else {
    var match = line.match(/^(\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2})\s+(.+)$/);
    if (match) {
      return {
        time: match[1].trim(),
        location: "",
        status: match[2].trim()
      };
    }
  }
  return null;
}
