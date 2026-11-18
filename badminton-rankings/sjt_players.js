/**
 * Finds all junior players within rank 50 from the "Ranking" sheet 
 * and outputs an aggregated summary to the "SJT Players" sheet.
 */
function extractSJTPlayers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rankingSheet = ss.getSheetByName("Ranking");
  
  if (!rankingSheet) {
    Logger.log("Error: 'Ranking' sheet not found.");
    return;
  }
  
  var data = rankingSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("Ranking sheet is empty or contains only headers.");
    return;
  }
  
  // Column Mapping based on your Ranking sheet structure:
  // 0: Age Group, 1: Event, 2: Placering, 3: Rank Change, 4: Spelare, 
  // 5: Licnr, 6: Födelseår, 7: Poäng, etc.
  
var playerMap = {}; // Key: Licnr, Value: Player Object with aggregated rankings
  var eventsSet = []; // Tracks distinct event categories dynamically (e.g., HS, DS, HD, DD, MD)
  
  // Loop through rows (skipping header row 0)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var ageGroup  = row[0];
    var event     = row[1];
    var placering = Number(row[2]);
    var name      = row[4];
    var licnr     = row[5];
    var birthYear = row[6];
    
    // Filter Criteria:
    // 1. Must be a junior player (Age Group is NOT "Senior")
    // 2. Placering must be a valid number between 1 and 50
    // 3. Must have a valid license number
    if (ageGroup !== "Senior" && !isNaN(placering) && placering > 0 && placering <= 50 && licnr) {
      
      // Track distinct event categories in order of appearance
      if (eventsSet.indexOf(event) === -1) {
        eventsSet.push(event);
      }

      // If we haven't seen this player yet, initialize their record
      if (!playerMap[licnr]) {
        playerMap[licnr] = {
          ageGroup: ageGroup,
          name: name,
          licnr: licnr,
          birthYear: birthYear,
          rankings: {} // Key-value map of { eventName: rank }
        };
      }
      
      // Store ranking under the corresponding event key
      playerMap[licnr].rankings[event] = placering;
    }
  }
  
  // Build headers dynamically: Base fields + distinct Event columns
  var headers = ["Age Group", "Name", "License Number", "Birth Year"].concat(eventsSet);
  var outputData = [headers];
  
  // Convert our player map object into output rows
  for (var lic in playerMap) {
    var p = playerMap[lic];
    var rowData = [
      p.ageGroup,
      p.name,
      p.licnr,
      p.birthYear
    ];

    // Append ranking for each event column, or an empty string if unranked in that event
    for (var j = 0; j < eventsSet.length; j++) {
      var evt = eventsSet[j];
      rowData.push(p.rankings[evt] !== undefined ? p.rankings[evt] : "");
    }

    outputData.push(rowData);
  }
  
  // Write output to "SJT Players" sheet
  var sjtSheet = ss.getSheetByName("SJT Players");
  if (!sjtSheet) {
    Logger.log("'SJT Players' sheet not found. Creating a new one.");
    sjtSheet = ss.insertSheet("SJT Players");
  }
  
  sjtSheet.clear();
  
  if (outputData.length >  1) {
    sjtSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    Logger.log("Success! Extracted " + (outputData.length - 1) + " unique junior players into 'SJT Players'.");
  } else {
    Logger.log("No junior players found within rank 50.");
  }
}