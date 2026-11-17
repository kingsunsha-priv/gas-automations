function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("51Ship")
    .addItem("Update transportation details", "UpdateTransportationDetails")
    .addToUi();
}

var globalSheet;
var globalData;

function GetSheet() {
  if (!globalSheet) {
    globalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Deliveries");
  }
  return globalSheet;
}

function GetData() {
  if (!globalData) {
    globalData = GetSheet().getDataRange().getValues();
  }
  return globalData;
}

function ClearDataCache() {
  globalData = null;
}

/**
 * Processes a single row in memory: fetches API tracking, updates country, and calculates ETA
 */
function ProcessShipmentRow(rowValues, headerMap, fullData) {
  var idCol = headerMap["运单号"];
  var typeCol = headerMap["种类"];
  var recCol = headerMap["已到货"];
  var errCol = headerMap["错误信息"];
  var countryCol = headerMap["目的国家"];
  var lastTrackCol = headerMap["末条追踪轨迹"];
  var detailLogCol = headerMap["详细轨迹"];
  
  // Extend to support both "ETA" and "预计到达" columns explicitly
  var etaCol = headerMap["ETA"] !== undefined ? headerMap["ETA"] : headerMap["预计到达"];

  if (idCol === undefined) return;

  var transId = String(rowValues[idCol]).trim();
  var transType = typeCol !== undefined ? String(rowValues[typeCol]).trim() : null;

  if (!transId) return;

  var detail = GetDetail(transId);

  if (detail.error !== "") {
    if (errCol !== undefined) rowValues[errCol] = detail.error;
    if (countryCol !== undefined) rowValues[countryCol] = "";
    if (lastTrackCol !== undefined) rowValues[lastTrackCol] = "";
    if (detailLogCol !== undefined) rowValues[detailLogCol] = "";
    if (etaCol !== undefined) rowValues[etaCol] = ""; // Explicitly clear ETA on error
  } else {
    if (errCol !== undefined) rowValues[errCol] = "";
    if (countryCol !== undefined) rowValues[countryCol] = NormalizeCountry(detail.country);
    if (lastTrackCol !== undefined) rowValues[lastTrackCol] = detail.lastTrackingInfo;
    if (detailLogCol !== undefined) rowValues[detailLogCol] = detail.eventLog;

    var isDelivered = detail.lastTrackingInfo.toUpperCase().indexOf("DELIVERED") >= 0 || detail.lastTrackingInfo === "已签收";
    if (recCol !== undefined) rowValues[recCol] = isDelivered;

    if (isDelivered) {
      // Explicitly clear ETA if the package is already delivered
      if (etaCol !== undefined) rowValues[etaCol] = ""; 
    } else {
      // Calculate ETA if package is active
      var hasETA = false;
      if (detail.eventLog && transType) {
        var lines = detail.eventLog.split('\n');
        if (lines.length > 0) {
          var parsed = ParseLogLine(lines[0]);
          if (parsed && parsed.status.toUpperCase().indexOf("DELIVERED") === -1) {
            var result = CalculateETA(transType, lines, detail.country, fullData, headerMap);
            if (result) {
              hasETA = true;
              if (lastTrackCol !== undefined) {
                rowValues[lastTrackCol] = detail.lastTrackingInfo + "\n" + result.logText;
              }
              if (etaCol !== undefined) {
                rowValues[etaCol] = result.etaDate; // Explicitly save the ETA date
              }
            }
          }
        }
      }
      
      // Explicitly clear ETA if calculation yielded no result
      if (!hasETA && etaCol !== undefined) {
        rowValues[etaCol] = "";
      }
    }
  }
}

/**
 * Main function to update all undelivered shipments in batch
 */
function UpdateTransportationDetails() {
  var sheet = GetSheet();
  var range = sheet.getDataRange();
  var dataValues = range.getValues();
  if (dataValues.length <= 1) return;

  var headerMap = buildHeaderMap(dataValues[0]);
  var recCol = headerMap["已到货"];
  var idCol = headerMap["运单号"];

  if (idCol === undefined || recCol === undefined) return;

  var hasModified = false;

  for (var i = 1; i < dataValues.length; i++) {
    var row = dataValues[i];
    var transId = row[idCol];
    var isReceived = row[recCol] === true || String(row[recCol]).toUpperCase() === "TRUE";

    if (transId && !isReceived) {
      ProcessShipmentRow(row, headerMap, dataValues);
      hasModified = true;
    }
  }

  // Single batch write back to Google Sheets
  if (hasModified) {
    range.setValues(dataValues);
    ClearDataCache();
  }
}

/**
 * ETA Engine with optimized single-pass pre-indexing, 1-year cutoff, and line-by-line cascading fallback
 */
function CalculateETA(currentType, currentLogLines, currentCountry, fullData, headerMap) {
  if (!currentType || !currentLogLines || currentLogLines.length === 0) return null;

  fullData = fullData || GetData();
  headerMap = headerMap || buildHeaderMap(fullData[0]);

  var typeCol = headerMap["种类"];
  var receivedCol = headerMap["已到货"];
  var logCol = headerMap["详细轨迹"];
  var idCol = headerMap["运单号"];
  var countryCol = headerMap["目的国家"];
  var shipDateCol = headerMap["发货日期"];

  if (typeCol === undefined || receivedCol === undefined || logCol === undefined || idCol === undefined) return null;

  var targetTypePrefix = String(currentType).substring(0, 2);
  var targetCountry = NormalizeCountry(currentCountry);

  // Define the cutoff date (1 year ago from today)
  var oneYearAgo = new Date();
  oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);

  // Pre-parse and index delivered historical shipments matching category & country
  var matchedHistories = [];
  for (var i = 1; i < fullData.length; i++) {
    var row = fullData[i];
    var isDelivered = row[receivedCol] === true || String(row[receivedCol]).toUpperCase() === "TRUE";
    var histTypePrefix = String(row[typeCol]).substring(0, 2);
    var histCountry = countryCol !== undefined ? NormalizeCountry(row[countryCol]) : "";

    var countryMatches = !targetCountry || !histCountry || (histCountry === targetCountry);

    if (isDelivered && histTypePrefix === targetTypePrefix && countryMatches) {
      var histLog = row[logCol];
      if (!histLog) continue;

      var histLines = String(histLog).split('\n');
      var delDate = null, delLine = "";
      var events = [];

      for (var j = 0; j < histLines.length; j++) {
        var parsed = ParseLogLine(histLines[j]);
        if (!parsed) continue;

        if (!delDate && (parsed.status.toUpperCase().indexOf("DELIVERED") >= 0 || parsed.status === "已签收")) {
          delDate = new Date(parsed.time.replace(/-/g, "/"));
          delLine = histLines[j];
        }
        events.push({ time: new Date(parsed.time.replace(/-/g, "/")), status: parsed.status, line: histLines[j] });
      }

      // ONLY include this historical record if it was delivered within the last 1 year
      if (delDate && delDate.getTime() >= oneYearAgo.getTime()) {
        var rawShipDate = shipDateCol !== undefined ? row[shipDateCol] : "";
        var formattedShipDate = "";
        if (rawShipDate instanceof Date) {
          formattedShipDate = Utilities.formatDate(rawShipDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (rawShipDate) {
          formattedShipDate = String(rawShipDate).substring(0, 10);
        }

        matchedHistories.push({
          id: row[idCol],
          shipDate: formattedShipDate || "未知发货日",
          delDate: delDate,
          delLine: delLine,
          events: events
        });
      }
    }
  }

  if (matchedHistories.length === 0) return null;

// Search logic across pre-indexed historical items
  function searchHistories(searchType, searchValue, logLineForBaseTime) {
    var parsedBase = ParseLogLine(logLineForBaseTime);
    if (!parsedBase) return null;

    var baseDate = new Date(parsedBase.time.replace(/-/g, "/"));
    var matchedDetails = [];

    for (var k = 0; k < matchedHistories.length; k++) {
      var hist = matchedHistories[k];
      var matchDate = null;
      var matchLine = "";

      for (var e = 0; e < hist.events.length; e++) {
        var ev = hist.events[e];
        
        // 新增 LOCATION 匹配逻辑
        if (searchType === "EXACT" && ev.status === searchValue) {
          matchDate = ev.time;
          matchLine = ev.line;
          break;
        } else if (searchType === "LOCATION" && ev.line.toUpperCase().indexOf(searchValue.toUpperCase()) >= 0) {
          matchDate = ev.time;
          matchLine = ev.line;
          break;
        } else if (searchType === "COUNTRY" && ev.line.toUpperCase().indexOf(searchValue.toUpperCase()) >= 0) {
          matchDate = ev.time;
          matchLine = ev.line;
          break;
        }
      }

      if (matchDate && hist.delDate) {
        var diff = hist.delDate.getTime() - matchDate.getTime();
        if (diff > 0) {
          matchedDetails.push({ 
            id: hist.id, 
            diff: diff,
            days: (diff / (1000 * 60 * 60 * 24)).toFixed(1), 
            matchLine: matchLine,
            delDate: hist.delDate
          });
        }
      }
    }

    if (matchedDetails.length > 0) {
      matchedDetails.sort(function(a, b) {
        return b.delDate.getTime() - a.delDate.getTime();
      });

      var topMatches = matchedDetails.slice(0, 3);
      var totalDiff = 0;
      for (var d = 0; d < topMatches.length; d++) {
        totalDiff += topMatches[d].diff;
      }

      var avgDiff = totalDiff / topMatches.length;
      var etaDate = new Date(baseDate.getTime() + avgDiff);
      var dateOnly = Utilities.formatDate(etaDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      var avgDays = (avgDiff / (1000 * 60 * 60 * 24)).toFixed(1); 

      var logText = "\n[基于最近 " + topMatches.length + " 个历史运单平均预测 ETA: " + dateOnly + " (平均需 " + avgDays + " 天)]\n";
      logText += "参考明细 (匹配节点 -> 签收耗时):\n";
      
      // only log the first match
      // for (var i = 0; i < topMatches.length; i++) {
      var i = 0;
      var detail = topMatches[i];
      var cleanMatchLine = detail.matchLine.split('\t').join(' | '); 
      
      logText += "- " + detail.id + " : 耗时 " + detail.days + " 天\n";
      // 更新了错别字：匹 -> 匹配
      logText += "  匹配: " + cleanMatchLine + "\n"; 
      // }

      return { etaDate: dateOnly, logText: logText.trimEnd() };
    }
    return null;
  }

  function getCountry(line) {
    if (!line) return null;
    var countries = ["Russia", "Kazakhstan", "Belarus", "Poland", "Czechia", "Germany", "Denmark", "Sweden", "China"];
    for (var c = 0; c < countries.length; c++) {
      if (line.toUpperCase().indexOf(countries[c].toUpperCase()) >= 0) return countries[c];
    }
    return null;
  }

  var result = null;
  var maxLogsToCheck = Math.min(3, currentLogLines.length);

  // Line-by-line evaluation: Exact Status -> Exact Location -> Generic Country
  for (var m = 0; m < maxLogsToCheck; m++) {
    var currentLine = currentLogLines[m];
    var parsedCurrent = ParseLogLine(currentLine);
    
    if (parsedCurrent) {
      // 1. 最高优先级：尝试 EXACT 状态匹配 (Status == Status)
      if (parsedCurrent.status) {
        result = searchHistories("EXACT", parsedCurrent.status, currentLine);
      }
      
      // 2. 次高优先级：尝试 EXACT 站点匹配 (Location == Location)
      if (!result && parsedCurrent.location) {
        result = searchHistories("LOCATION", parsedCurrent.location, currentLine);
      }
    }
    
    // 3. 最低优先级：兜底尝试 COUNTRY 宽泛匹配
    if (!result) {
      var cMatch = getCountry(currentLine);
      if (cMatch) {
        result = searchHistories("COUNTRY", cMatch, currentLine);
      }
    }
    
    // 如果找到了匹配项（不论是状态、站点还是国家），立即停止，不再往下找
    if (result) break; 
  }

  return result;
}

function GetDetail(transId) {
  Logger.log("Process transId: " + transId);
  try {
    var response = UrlFetchApp.fetch("http://47.103.28.217/public/v1/ordeSearch/trackInfoList?number=" + transId).getContentText();
    var detail = JSON.parse(response);

    if (detail.rows && detail.rows.length > 0) {
      var row = detail.rows[0];
      var events = row.trackEventsList || [];
      
      // Force sort descending by time (Newest First)
      events.sort(function(a, b) {
        var timeA = new Date((a.createTime || "").replace(/-/g, "/")).getTime() || 0;
        var timeB = new Date((b.createTime || "").replace(/-/g, "/")).getTime() || 0;
        return timeB - timeA;
      });

      var eventLog = "";
      for (var e = 0; e < events.length; e++) {
        if (eventLog !== "") eventLog += "\n";
        eventLog += events[e].createTime + (events[e].location === "" ? "" : ("\t" + events[e].location)) + "\t" + events[e].description;
      }

      return {
        error: "",
        country: row.country,
        lastTrackingInfo: row.lastTrackInfo,
        eventLog: eventLog
      };
    }

    return GetTrainDetail(transId);
  } catch (ex) {
    return { error: "Error: " + ex.message };
  }
}

function GetTrainDetail(transId) {
  if (!transId.startsWith("1Z")) {
    return { error: "Not a recognized UPS train tracking number." };
  }

  var API_KEY_17TRACK = "E241A7663DA380846453FC364A8EA24E";
  var proxyUrl = "https://www.51ship.cc/tracking/ups?trackingNo=" + transId;
  var proxyOptions = {
    muteHttpExceptions: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
      "Accept": "application/json",
      "Referer": "https://www.51ship.cc/"
    }
  };

  try {
    var proxyResponse = UrlFetchApp.fetch(proxyUrl, proxyOptions);
    var proxyText = proxyResponse.getContentText().trim();

    if (proxyText && proxyText.indexOf("[") === 0) {
      var data = JSON.parse(proxyText);
      if (Array.isArray(data) && data.length > 0) {
        
        var validScans = [];
        for (var i = 0; i < data.length; i++) {
          if (data[i].status.type !== 'M') { // Ignore manifest
            var dateRaw = data[i].date + data[i].time;
            var formattedDate = dateRaw.length === 14 
              ? dateRaw.substr(0,4) + '-' + dateRaw.substr(4,2) + '-' + dateRaw.substr(6,2) + ' ' + dateRaw.substr(8,2) + ':' + dateRaw.substr(10,2)
              : dateRaw;
              
            validScans.push({
              sortKey: dateRaw, // YYYYMMDDHHMMSS formats perfectly for string sorting
              formattedDate: formattedDate,
              location: data[i].location.address.city + ", " + data[i].location.address.country,
              description: data[i].status.description,
              country: data[i].location.address.country
            });
          }
        }

        // Force sort descending by time string (Newest First)
        validScans.sort(function(a, b) {
          return b.sortKey.localeCompare(a.sortKey);
        });

        if (validScans.length > 0) {
          var eventLog = "";
          for (var k = 0; k < validScans.length; k++) {
            if (eventLog !== "") eventLog += "\n";
            eventLog += validScans[k].formattedDate + "\t" + validScans[k].location + "\t" + validScans[k].description;
          }

          return {
            error: "",
            country: validScans[0].country || "EU", // Pull country from newest physical scan
            lastTrackingInfo: "UPS Handling Package",
            eventLog: eventLog
          };
        }
      }
    }

    // Fallback: 17Track API
    if (API_KEY_17TRACK !== "YOUR_17TRACK_API_KEY") {
      var trackPayload = [{ "number": transId }];
      var trackOptions = {
        "method": "post",
        "contentType": "application/json",
        "headers": { "17token": API_KEY_17TRACK },
        "payload": JSON.stringify(trackPayload),
        "muteHttpExceptions": true
      };

      UrlFetchApp.fetch("https://api.17track.net/track/v2.2/register", trackOptions);
      Utilities.sleep(2000);

      var trackResponse = UrlFetchApp.fetch("https://api.17track.net/track/v2.2/gettrackinfo", trackOptions);
      var trackData = JSON.parse(trackResponse.getContentText());

      if (trackData.code === 0 && trackData.data && trackData.data.accepted && trackData.data.accepted.length > 0) {
        var packageInfo = trackData.data.accepted[0].track_info;
        if (packageInfo && packageInfo.tracking && packageInfo.tracking.providers && packageInfo.tracking.providers.length > 0) {
          var fbEventLog = "";
          
          if (packageInfo.time_metrics && packageInfo.time_metrics.estimated_delivery_date) {
            var est = packageInfo.time_metrics.estimated_delivery_date;
            var targetDate = est.to || est.from;
            if (targetDate) {
              fbEventLog += "[Estimated Delivery: " + targetDate.substring(0, 10) + "]\n";
            }
          }

          var events = packageInfo.tracking.providers[0].events;
          
          // Force sort descending by ISO time string (Newest First)
          events.sort(function(a, b) {
            var timeA = new Date((a.time_iso || "").replace("T", " ")).getTime() || 0;
            var timeB = new Date((b.time_iso || "").replace("T", " ")).getTime() || 0;
            return timeB - timeA;
          });

          // Loop forward normally since the array is now explicitly sorted newest-to-oldest
          for (var j = 0; j < events.length; j++) {
            if (fbEventLog !== "") fbEventLog += "\n";
            var rawTime = events[j].time_iso || "";
            var cleanTime = rawTime ? rawTime.replace("T", " ").substring(0, 19) : "";
            var loc = events[j].location || "EU";
            var desc = events[j].description || "";
            fbEventLog += cleanTime + "\t" + loc + "\t" + desc;
          }

          var currentStatus = "EU Scan";
          if (packageInfo.latest_status && packageInfo.latest_status.status) {
            currentStatus = packageInfo.latest_status.status;
            if (packageInfo.latest_status.sub_status) {
              currentStatus += " (" + packageInfo.latest_status.sub_status + ")";
            }
          }

          return {
            error: "",
            country: packageInfo.shipping_info ? (packageInfo.shipping_info.recipient_address.country || "EU") : "EU",
            lastTrackingInfo: currentStatus,
            eventLog: fbEventLog
          };
        }
      }
    }

    return {
      error: "",
      country: "Transit / No Data",
      lastTrackingInfo: "Awaiting EU Scan / Archived",
      eventLog: "No live scanning events found on 51ship or 17Track. The shipment is either currently in bulk transit on the train, or was delivered far enough in the past that carrier logs were purged."
    };
  } catch (ex) {
    return { error: "Error during fetch: " + ex.message };
  }
}
