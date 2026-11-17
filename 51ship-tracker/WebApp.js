function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('51Ship 物流追踪')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) 
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getShipmentData() {
  ClearDataCache();
  var sheet = GetSheet();
  
  // 对于团号，我们只需要用到普通的文本值即可，避开复杂的日期解析
  var displayValues = sheet.getDataRange().getDisplayValues(); 
  
  if (displayValues.length <= 1) {
    return displayValues;
  }

  var headers = displayValues[0];
  var groupCol = headers.indexOf("团号");
  var rows = displayValues.slice(1);

  if (groupCol !== -1) {
    rows.sort(function(a, b) {
      var valA = String(a[groupCol] || "").trim();
      var valB = String(b[groupCol] || "").trim();
      
      // 1. 把没有团号的空白行强制沉到最底部
      if (valA === "" && valB !== "") return 1;
      if (valB === "" && valA !== "") return -1;
      if (valA === valB) return 0;
      
      // 2. 尝试转换为数字进行严谨排序（如果团号是纯数字或日期如 20261107）
      var numA = Number(valA);
      var numB = Number(valB);
      
      if (!isNaN(numA) && !isNaN(numB)) {
        return numB - numA; // 纯数字降序 (大数在前)
      }
      
      // 3. 如果团号包含字母（如 T261107），则回退到标准的字符串降序排列
      return valA < valB ? 1 : -1;
    });
  }

  // 将排序后的数据与表头重新拼合并返回给前端
  return [headers].concat(rows);
}

function refreshAllShipments() {
  UpdateTransportationDetails();
  return getShipmentData();
}

function addShipmentAndRefresh(trackingNo, transType, groupNo, groupName, shipDate) {
  var sheet = GetSheet();
  var fullData = sheet.getDataRange().getValues();
  var headers = fullData[0];
  var headerMap = buildHeaderMap(headers);

  var newTrack = String(trackingNo).trim();

  // 1. Check for duplicates before doing anything else
  var trackCol = headerMap["运单号"];
  if (trackCol !== undefined) {
    for (var i = 1; i < fullData.length; i++) {
      if (String(fullData[i][trackCol]).trim() === newTrack) {
        throw new Error("该运单号已存在，无法重复添加！(Duplicate Tracking Number: " + newTrack + ")");
      }
    }
  }

  // 2. Build the new row array
  var newRow = new Array(headers.length).fill("");
  
  if (headerMap["运单号"] !== undefined) newRow[headerMap["运单号"]] = newTrack;
  if (headerMap["种类"] !== undefined) newRow[headerMap["种类"]] = String(transType).trim();
  if (headerMap["已到货"] !== undefined) newRow[headerMap["已到货"]] = false;
  if (headerMap["团号"] !== undefined) newRow[headerMap["团号"]] = String(groupNo).trim();
  if (headerMap["团名"] !== undefined) newRow[headerMap["团名"]] = String(groupName).trim();
  
  if (headerMap["发货日期"] !== undefined) {
    newRow[headerMap["发货日期"]] = String(shipDate).trim();
  }

  // 3. Process tracking details for the new row in memory
  ProcessShipmentRow(newRow, headerMap, fullData);

  // 4. Append new row directly
  sheet.appendRow(newRow);
  ClearDataCache();

  return getShipmentData();
}

function updateShipment(trackingNo, groupName, transType, shipDate) {
  var sheet = GetSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var trackCol = headers.indexOf("运单号");
  var groupNameCol = headers.indexOf("团名");
  var typeCol = headers.indexOf("种类");
  // Check for either date column name to be safe
  var dateCol = headers.indexOf("发货日期");
  
  if (trackCol === -1) return getShipmentData();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][trackCol]).trim() === String(trackingNo).trim()) {
      // Overwrite only the allowed fields
      if (groupNameCol !== -1) sheet.getRange(i + 1, groupNameCol + 1).setValue(groupName);
      if (typeCol !== -1) sheet.getRange(i + 1, typeCol + 1).setValue(transType);
      if (dateCol !== -1) sheet.getRange(i + 1, dateCol + 1).setValue(shipDate);
      
      ClearDataCache();
      break;
    }
  }
  
  return getShipmentData();
}
