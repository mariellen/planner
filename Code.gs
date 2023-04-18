/* todo
cost at date/sprint for each initiative
cost per sprint
cost per sprint per initiative
*/



let magic_totalInSchedule = "Total in Schedule";
let magic_currentEstimates = "Current Estimate";
let magic_Differences = "Difference";
let magic_Sprints = "Sprints";
let magic_Initiatives = "Initiative";
let magic_Background = "Background";
let magic_DevNamesBelow = "Dev Names Below";

let index_outputInitiatives = 1;
let index_outputCurrentEstimates = 2;
let index_outputTotalInShedule = 3;
let index_outputDifferences = 4;
let index_outputSprints = 5;
let index_lookupDevNames = 1;
let index_lookupDevLocation = 2;
let index_lookupDevCost = 3;
let index_lookupInitiatives = 6;
let index_devNamesBelow = 8;
let sheetName_lookup = "Lookups";

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Do Le Calculationz")
    .addItem("calculate days per initiative", "calculateDaysPerInitiative")
    .addItem("calculate days per initiative per dev", "calculateEverything")
    .addToUi();
}

function getSprintNames() {
  let scheduleSheet = SpreadsheetApp.getActive().getSheetByName("Sprint Header");
  let scheduleRange = scheduleSheet.getRange("Schedule!A1:Z30");
  let lastRow = scheduleRange.getLastRow();
  let lastColumn = scheduleRange.getLastColumn();
  let sprintNames = Array();
  let sprintNameIndexes = Array();
  let sprintNameRow = 0;

  for (let r = 1; r < lastRow - 1; r++) {
    for (let c = 1; c < lastColumn - 1; c++) {
      let cell = scheduleRange.getCell(r, c);
      let cellValue = cell.getValue();
      if (c == 1 && cellValue.toLowerCase() == "sprint name") {
        sprintNameRow = r;
      }
      if (r == sprintNameRow && c > 1 && cellValue != "") {
        sprintNames.push(cellValue);
        sprintNameIndexes[c] = cellValue;
      }
    }
  }
  let sprintObj = new Array();
  sprintObj["sprintNames"] = sprintNames;
  sprintObj["sprintNameIndexes"] = sprintNameIndexes;
  return sprintObj;
}

function calculateEverything() {
  calculateDaysPerInitiative(true);
}

function sumInitiativeByBackgroundColour(scheduleSheetName, backgroundColour) {
  const refSheet = SpreadsheetApp.getActive().getSheetByName(scheduleSheetName);

  const startRow = index_devNamesBelow + 1; // configurable start row
  const startCol = 2; // configurable start column
  const lastRow = refSheet.getLastRow();
  const lastCol = refSheet.getLastColumn();

  const range = sheet.getRange(startRow, startCol, lastRow - startRow + 1, lastCol - startCol + 1);
  let sumOfColouredCells = 0;
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[0].length; c++) {
      const cellValue = values[r][c];
      const cellBackground = backgrounds[r][c];
      if (cellValue !== "" && cellBackground === backgroundColour) {
        sumOfColouredCells += cellValue;
      }
    }
}

  return sumOfColouredCells;
}

function SumInitiative(scheduleSheetName, row, col) {
  let backgroundColour = getBackgroundColourForInitiative(row, col);
  return SumInitiativeByBackgroundColour(scheduleSheetName, backgroundColour);
}

function UpdateSumInitiative(scheduleSheetName, row, col) {
  let a = SumInitiative(scheduleSheetName, row, col);
  SpreadsheetApp.getActiveSheet().getRange(2, 3).setValue(a);
}

function getBackgroundColourForInitiative(row, col) {
  let backgroundColour = SpreadsheetApp.getActiveSheet().getRange(row, col).getBackground();
  return backgroundColour;
}

function getValuesAs2DArray(things) {
  let thingArray = Array();
  for (r = 0; r < things.length; r++) {
    thingArray[r] = Array();
    thingArray[r][0] = things[r];
  }
  return thingArray;
}

function getInitiatives() {
  const initiatives = [];
  const estimates = [];
  const bgColours = [];
  let initiativeColumn = 0;
  let bgColumn = 0;
  let estimatesColumn = 0;
  const lookupsSheet = SpreadsheetApp.getActive().getSheetByName(sheetName_lookup);
  const lookupsDataRange = lookupsSheet.getDataRange();
  const lookupsValues = lookupsDataRange.getValues();
  const columnValues = lookupsValues[0];
  columnValues.forEach((value, index) => {
    if (value === magic_Initiatives) {
      initiativeColumn = index + 1;
    }
    if (value === magic_currentEstimates) {
      estimatesColumn = index + 1;
    }
    if (value === magic_Background) {
      bgColumn = index + 1;
    }
  });
  if (initiativeColumn > 0 && estimatesColumn > 0 && bgColumn > 0) {
    for (let ro = 1; ro < lookupsValues.length; ro++) {
      initiatives.push(lookupsValues[ro][initiativeColumn - 1]);
      estimates.push(lookupsValues[ro][estimatesColumn - 1]);
      bgColours.push(lookupsValues[ro][bgColumn - 1]);
    }
  }
  return {
    initiatives,
    estimates,
    bgColours
  };
}

function applyConditionalFormatting(columnIndex, range, sheet) {
  const columnLetter = getColumnLetter(columnIndex);

  // Rule 1: Value less than 20% or greater than 40% is highlighted with a dark red background
  const rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=OR(${columnLetter}:${columnLetter}<20%,${columnLetter}:${columnLetter}>40%)`)
    .setBackground("#8b0000")
    .setFontColor("red")
    .setBold(true)
    .setRanges([range])
    .build();

  // Rule 2: Value between 0% and 20% is highlighted with an orange background
  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(${columnLetter}:${columnLetter}>=0%,${columnLetter}:${columnLetter}<20%)`)
    .setBackground("#ff8c00")
    .setFontColor("red")
    .setRanges([range])
    .build();

  // Rule 3: Value between 20% and 40% is highlighted with a green background
  const rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(${columnLetter}:${columnLetter}>=20%,${columnLetter}:${columnLetter}<=40%)`)
    .setBackground("#008000")
    .setRanges([range])
    .build();

  // Apply the conditional formatting rules to the sheet
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule3, rule2, rule1); // Add rules in reverse order to ensure correct priority
  sheet.setConditionalFormatRules(rules);
}

// thank you chatGPT
function getColumnLetter(columnIndex) {
  var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var letter = "";

  if (columnIndex <= 0) {
    throw new Error("Invalid column index");
  }

  while (columnIndex > 0) {
    var remainder = (columnIndex - 1) % 26;
    letter = alphabet.charAt(remainder) + letter;
    columnIndex = Math.floor((columnIndex - 1) / 26);
  }

  return letter;
}

// thank you chatGPT
function arrayToCommaDelimitedList(arr) {
  // Remove duplicates from the array
  arr = Array.from(new Set(arr));

  // Sort the array in alphabetical order
  arr.sort();

  // Filter out empty items from the array
  arr = arr.filter(function (item) {
    return item !== "";
  });

  // Join the filtered and sorted array into a comma-delimited list
  var commaDelimitedList = arr.join(", ");

  // Remove trailing commas
  commaDelimitedList = commaDelimitedList.replace(/,\s*$/, "");

  return commaDelimitedList;
}

function calculateDaysOnlyWithRecreate() {
  calculateDaysPerInitiative(false, true);
}

function calculateDaysPerInitiative(doCosts = false, rebuildSheet = false) {

  let scheduleSheetName = SpreadsheetApp.getActiveSheet().getName();
  let summarySheetName = scheduleSheetName + "Summary";

  // Use regular expression to check if sheet name starts with "Schedule" and does not end with "Summary"
  var pattern = /^schedule(?!.*summary$)/i; // i flag for case-insensitive matching
  if (!scheduleSheetName.match(pattern)) {
    Browser.msgBox("Error", "Your request could not be completed, run the command from a sheet with a name starting with 'Schedule' and not a Summary sheet.", Browser.Buttons.OK);

    return;
  }

  // Open the sheet by its name
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_lookup);
  var initiativesObj = getInitiatives();

  // Get the values in column A
  var range = sheet.getRange("A" + index_lookupDevNames + ":A");
  var values = range.getValues();

  // Loop through the values from bottom to top to find the last non-empty row
  var allDevCount = range.getLastRow();
  for (var i = allDevCount - 1; i >= 0; i--) {
    if (values[i][0] != "") {
      allDevCount = i + 1;
      break;
    }
  }

  let allDevs = SpreadsheetApp.getActive().getSheetByName(sheetName_lookup).getRange(2, index_lookupDevNames, allDevCount, 3);
  let summarySheet = SpreadsheetApp.getActive().getSheetByName(summarySheetName);

  if (summarySheet != null && rebuildSheet == true) {
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(summarySheet);
    summarySheet = null;
  }

  if (summarySheet == null) {
    summarySheet = createSummarySheet(summarySheetName, initiativesObj);
  }
  else {
    summarySheet.getRange(1, 5, summarySheet.getLastRow(), summarySheet.getLastColumn()).clearContent();
    updateSummarySheet(summarySheetName, initiativesObj);
  }

  //Browser.msgBox(4);
  let scheduleSheet = SpreadsheetApp.getActive().getSheetByName(scheduleSheetName);
  let scheduleRange = scheduleSheet.getRange(index_devNamesBelow, 1, allDevCount, scheduleSheet.getDataRange().getLastColumn());
  let lastScheduleRow = scheduleRange.getLastRow();
  let lastScheduleColumn = scheduleRange.getLastColumn();
  let devNames = Array();
  let differences = Array();
  let scheduledDays = Array();
  let sumOfDaysByBackgroundColour = Array();
  let devNameIndexes = Array();
  let sprintNamesByInitiativeColour = Array();
  let sprintObj = getSprintNames();
  let sprintNameByColumn = sprintObj["sprintNameIndexes"];
  var devCount = 0;
  for (let r = 1; r <= allDevCount; r++) {
    for (let c = 1; c < lastScheduleColumn; c++) {

      let thisSprintName = sprintNameByColumn[c];
      let cell = scheduleRange.getCell(r, c);
      let cellBackground = cell.getBackground();
      let cellValue = cell.getValue();
      if (cellValue != null && cellValue !== undefined && cellValue !== "") {

        if (r == 1 && c == 1 && cellValue != magic_DevNamesBelow) {
          Browser.msgBox("Your devs should be in row " + index_devNamesBelow + " onwards, with the heading in row " + index_devNamesBelow + " being '" + magic_DevNamesBelow + "'.");
        }

        if (c == 1) {
          let devName = cellValue;
          if (!Object.keys(devNames).some(key => key === devName)) {
            let dev = Array();
            dev["Name"] = devName;
            dev["Initiatives"] = Array();
            dev["Total"] = 0;
            devNames[r] = dev;
            devNameIndexes[devName] = r;
            devCount++;
          }
        }

        if (cellValue != "" && !isNaN(cellValue)) {
          if (!Object.keys(sumOfDaysByBackgroundColour).some(key => key === cellBackground)) {
            sumOfDaysByBackgroundColour[cellBackground] = 0;
          }
          sumOfDaysByBackgroundColour[cellBackground] += cellValue;

          if (thisSprintName != "") {
            if (!Object.keys(sprintNamesByInitiativeColour).some(key => key === cellBackground)) {
              sprintNamesByInitiativeColour[cellBackground] = Array();
            }
            sprintNamesByInitiativeColour[cellBackground].push(thisSprintName);
          }

          let assignedDev = devNames[r];
          if (!Object.keys(assignedDev["Initiatives"]).some(key => key === cellBackground)) {
            assignedDev["Initiatives"][cellBackground] = 0;
          }
          assignedDev["Initiatives"][cellBackground] += cellValue;
          assignedDev["Total"] += cellValue;
        }
      }
    }
  }

  let concatSprintNamesPerInitiative = Array();

  let summarySheetLastRow = summarySheet.getLastRow();
  for (let ro = 2; ro <= summarySheetLastRow; ro++) {
    let cell = summarySheet.getRange(ro, 1);
    let backgroundColour = cell.getBackground();
    let sumOfDays = sumOfDaysByBackgroundColour[backgroundColour];
    scheduledDays.push(sumOfDays);

    let estimateCell = summarySheet.getRange(ro, index_outputCurrentEstimates);
    let estimate = estimateCell.getValue();
    let difference = ((sumOfDays - estimate) / sumOfDays);
    differences.push(difference);

    let sprintsForThisInitiative = sprintNamesByInitiativeColour[backgroundColour];
    if (sprintsForThisInitiative != null) {
      uniq = arrayToCommaDelimitedList(sprintsForThisInitiative);
      concatSprintNamesPerInitiative.push(uniq);
    }
  }

  summarySheet.getRange(2, index_outputTotalInShedule, scheduledDays.length, 1).setValues(getValuesAs2DArray(scheduledDays));
  summarySheet.getRange(2, index_outputDifferences, differences.length, 1).setValues(getValuesAs2DArray(differences));
  summarySheet.getRange(2, index_outputSprints, concatSprintNamesPerInitiative.length, 1).setValues(getValuesAs2DArray(concatSprintNamesPerInitiative));

  let summarySheetLastColumn = summarySheet.getLastColumn();

  if (doCosts == true) {


    // for each dev in the schedule
    for (var co = 1; co < allDevs.getLastRow(); co++) {
      var devToInitiative = Array();
      var cell = allDevs.getCell(co, 1);
      var devName = cell.getValue();
      var costCell = allDevs.getCell(co, 3);
      var cost = costCell.getValue();

      if (devName !== "") {
        devToInitiative.push(devName);
        var devObjectIndex = devNameIndexes[devName];
        var devObject = devNames[devObjectIndex];
        if (devObject != null) {

          // for each initiative in the output summary
          for (var ro = 2; ro <= summarySheetLastRow; ro++) {
            var cell = summarySheet.getRange(ro, index_outputInitiatives);
            var backgroundColour = cell.getBackground();
            var sumOfDaysThisInitiative = 0;

            if (Object.keys(devObject).some(key => key === "Initiatives") &&
              Object.keys(devObject["Initiatives"]).some(key => key === backgroundColour)) {
              sumOfDaysThisInitiative = devObject["Initiatives"][backgroundColour];
            }
            devToInitiative.push(sumOfDaysThisInitiative);
          }
          devToInitiative.push(devObject["Total"]);
          devToInitiative.push(devObject["Total"] * cost);
        }
        else {
          Browser.msgBox("There is no information for: " + devName);
        }
        summarySheet.getRange(1, summarySheetLastColumn + 1 + co, devToInitiative.length, 1).setValues(getValuesAs2DArray(devToInitiative));
      }
    }

    summarySheetLastColumn = summarySheet.getLastColumn();

    let internaldevToInitiative = Array();
    let externaldevToInitiative = Array();
    let externalcostToInitiative = Array();

    let totalExternalCost = 0;
    let totalInternalDays = 0;
    let totalExternalDays = 0;

    internaldevToInitiative.push("Internal Days");
    externaldevToInitiative.push("External Days");
    externalcostToInitiative.push("External Cost");

    // foreach initiative in output sheet
    for (let ro = 2; ro <= summarySheetLastRow; ro++) {
      let internalDaysThisInitiative = 0;
      let externalDaysThisInitiative = 0;
      let externalCostThisInitiative = 0;
      let cell = summarySheet.getRange(ro, index_outputInitiatives);
      let backgroundColour = cell.getBackground();

      // foreach dev in the lookup table
      for (let co = 1; co < allDevs.getLastRow(); co++) {
        let cell = allDevs.getCell(co, index_lookupDevNames);
        let devName = cell.getValue();
        let locationCell = allDevs.getCell(co, index_lookupDevLocation);
        let location = locationCell.getValue();
        let costCell = allDevs.getCell(co, index_lookupDevCost);
        let cost = costCell.getValue();

        let devObjectIndex = devNameIndexes[devName];
        let devObject = devNames[devObjectIndex];
        let sumOfDaysThisInitiative = 0;
        if (devObject != null &&
          Object.keys(devObject).some(key => key === "Initiatives") &&
          Object.keys(devObject["Initiatives"]).some(key => key === backgroundColour)) {
          sumOfDaysThisInitiative = devObject["Initiatives"][backgroundColour];
        }
        if (location == "Internal") {
          internalDaysThisInitiative += sumOfDaysThisInitiative;
          totalExternalCost += sumOfDaysThisInitiative;
          totalInternalDays += sumOfDaysThisInitiative;
        }
        else {
          externalDaysThisInitiative += sumOfDaysThisInitiative;
          externalCostThisInitiative += (sumOfDaysThisInitiative * cost);
          totalExternalDays += sumOfDaysThisInitiative;

        }
      }
      internaldevToInitiative.push(internalDaysThisInitiative);
      externaldevToInitiative.push(externalDaysThisInitiative);
      externalcostToInitiative.push(externalCostThisInitiative);
      totalExternalCost += externalCostThisInitiative;      
    }

    setFirstCellValue(scheduleSheetName, totalExternalCost);

    externalcostToInitiative.push("-");
    externalcostToInitiative.push(totalExternalCost);
    externaldevToInitiative.push(totalExternalDays);
    internaldevToInitiative.push(totalInternalDays);
    
    summarySheetLastColumn = summarySheet.getLastColumn();
    summarySheet.getRange(1, summarySheetLastColumn + 1, internaldevToInitiative.length, 1).setValues(getValuesAs2DArray(internaldevToInitiative));
    summarySheet.getRange(1, summarySheetLastColumn + 2, externaldevToInitiative.length, 1).setValues(getValuesAs2DArray(externaldevToInitiative));
    summarySheet.getRange(1, summarySheetLastColumn + 3, externalcostToInitiative.length, 1).setValues(getValuesAs2DArray(externalcostToInitiative));

    summarySheet.getRange(summarySheet.getLastRow(), 1, 1, summarySheet.getLastColumn()).setNumberFormat("$00.00");
    summarySheet.getRange(summarySheet.getLastRow(), 1, 1, summarySheet.getLastColumn()).setFontWeight("bold");
    summarySheet.getRange(1, summarySheet.getLastColumn(), summarySheet.getLastRow(), 1).setNumberFormat("$00.00");
    summarySheet.getRange(1, summarySheet.getLastColumn(), summarySheet.getLastRow(), 1).setFontWeight("bold");
  }

  summarySheet.autoResizeColumns(1, summarySheet.getLastColumn());
  var columnIndexes = [index_outputCurrentEstimates, index_outputTotalInShedule, index_outputDifferences];
  applyFormattingToSummarySheet(summarySheetName, initiativesObj, columnIndexes);

  let targetRow = allDevCount + index_devNamesBelow + 3;
  var sourceRange = summarySheet.getRange('A1:D' + initiativesObj.initiatives.length);
  let bottomRowToUse = targetRow + 4;
  transposeDataWithFormat(scheduleSheetName, targetRow, sourceRange);
  
  if (doCosts)
  {
    sourceRange = getLastThreeColumnsRange(summarySheetName, initiativesObj.initiatives.length + 2); // also get the totals
    transposeDataWithFormat(scheduleSheetName, bottomRowToUse, sourceRange);
    bottomRowToUse += 3;
    
  }
  scheduleSheet.getRange(bottomRowToUse, 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).clearContent();
  scheduleSheet.getRange(scheduleSheet.getDataRange().getLastRow(), 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).setFontWeight("bold");
}

function getLastThreeColumnsRange(sheetName, numRows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var numCols = sheet.getDataRange().getLastColumn();
  
  // Check if there are at least three columns in the sheet
  if (numCols < 3) {
    throw new Error('The sheet must have at least three columns.');
  }
  
  var startCol = numCols - 2;  
  var range = sheet.getRange(1, startCol, numRows, 3);
  return range;
}

function transposeDataWithFormat(scheduleSheetName, targetRow, sourceRange) {
  
  var currentSprintCol = getCurrentDateRangeColumn(scheduleSheetName);

  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(scheduleSheetName); // Get the target sheet by name

  var sourceData = sourceRange.getValues(); // Get data from range A:F in the source sheet
  var sourceNumberFormat = sourceRange.getNumberFormats(); // Get number formats from range A:F in the source sheet

  var transposedData = transposeArray(sourceData); // Transpose the data  
  var transposedNumberFormat = transposeArray(sourceNumberFormat); // Transpose the number formats

  var rowsOfData = transposedData.length;
  var colsOfData = transposedData[0].length;

  targetSheet.getRange(targetRow, 1, targetSheet.getLastRow(), targetSheet.getLastColumn()).clearContent();
  targetSheet.getRange(targetRow, 1, targetSheet.getLastRow(), targetSheet.getLastColumn()).clearFormat();

  // paste the summaries into the current sprint
  var targetRange = targetSheet.getRange(targetRow, currentSprintCol, rowsOfData, colsOfData);

  targetRange.setValues(transposedData); // Insert the transposed data into the target sheet
  targetRange.setNumberFormats(transposedNumberFormat); // Set the number formats in the target sheet
  targetRange.setBackgrounds(transposeArray(sourceRange.getBackgrounds()));
  targetRange.setFontColors(transposeArray(sourceRange.getFontColors()));

  var bandings = sourceRange.getBandings();
  if (bandings.length > 0) {
    var sourceFormat = bandings[0].getRange().getCell(1, 1).getRichTextValue().getTextStyle();
    var transposedFormat = transposeArray(sourceFormat); // Transpose the number formats
    targetRange.setFontFamilies(transposedFormat.getFontFamily());
    targetRange.setFontSizes(transposedFormat.getFontSize());
    targetRange.setFontStyles(transposedFormat.isItalic(), transposedFormat.isBold(), transposedFormat.isUnderline());
  }

}


function getColumnRange(sheetName, columnIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var numRows = sheet.getLastRow();
  var rangeA1Notation = "";
  var columnLetter = String.fromCharCode(65 + columnIndex - 1);
  rangeA1Notation += columnLetter + "1:" + columnLetter + numRows;
  return sheet.getRange(rangeA1Notation);
}


function transposeArray(array) {
  var newArray = [];
  for (var i = 0; i < array[0].length; i++) {
    newArray.push([]);
    for (var j = 0; j < array.length; j++) {
      newArray[i].push(array[j][i]);
    }
  }
  return newArray;
}

function createSummarySheet(summarySheetName, initiativesObj) {
  summarySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  summarySheet.setName(summarySheetName);
  const initiatives = initiativesObj.initiatives;
  const estimates = initiativesObj.estimates;
  summarySheet.getRange(1, index_outputInitiatives, 1, 1).setValue(magic_Initiatives);
  summarySheet.getRange(1, index_outputCurrentEstimates, 1, 1).setValue(magic_currentEstimates);
  summarySheet.getRange(1, index_outputTotalInShedule, 1, 1).setValue(magic_totalInSchedule);
  summarySheet.getRange(1, index_outputDifferences, 1, 1).setValue(magic_Differences);
  summarySheet.getRange(1, index_outputSprints, 1, 1).setValue(magic_Sprints);
  summarySheet.getRange(2, index_outputInitiatives, initiatives.length, 1).setValues(getValuesAs2DArray(initiatives));
  summarySheet.getRange(2, index_outputCurrentEstimates, estimates.length, 1).setValues(getValuesAs2DArray(estimates));
  var columnIndexes = [index_outputDifferences];
  applyFormattingToSummarySheet(summarySheetName, initiativesObj, columnIndexes);
  return summarySheet;
}

function updateSummarySheet(summarySheetName, initiativesObj) {
  summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheetName);
  const initiatives = initiativesObj.initiatives;
  const estimates = initiativesObj.estimates;
  summarySheet.getRange(2, index_outputInitiatives, initiatives.length, 1).setValues(getValuesAs2DArray(initiatives));
  summarySheet.getRange(2, index_outputCurrentEstimates, estimates.length, 1).setValues(getValuesAs2DArray(estimates));
  return summarySheet;
}

function applyFormattingToSummarySheet(summarySheetName, initiativesObj, columnIndexes) {
  var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheetName);
  const bgColours = initiativesObj.bgColours;

  for (let bg = 0; bg < bgColours.length; bg++) {
    summarySheet.getRange(bg + 2, index_outputInitiatives, 1, 1).setBackground(bgColours[bg]);
    summarySheet.getRange(bg + 2, index_outputDifferences, 1, 1).setNumberFormat('0.00%');
  }

  for (let i = 0; i < columnIndexes.length; i++) {
    var rangeA1Notation = getColumnRange(summarySheetName, columnIndexes[i]);
    applyConditionalFormatting(index_outputDifferences, rangeA1Notation, summarySheet);
  }

  return summarySheet;
}

function getCurrentDateRangeColumn(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  var currentDate = new Date();  // Get current date
  var startRow = 4;  // Start date is in row 4
  var endRow = 5;  // End date is in row 5
  
  var values = sheet.getRange(startRow, 1, endRow, sheet.getLastColumn()).getValues();
  
  for (var col = 0; col < values[0].length; col++) {
    var startDate = new Date(values[0][col]);
    var endDate = new Date(values[1][col]);
    
    if (currentDate >= startDate && currentDate <= endDate) {
      // Current date is within this date range, return the column index
      return col + 1;  // Adjust for 0-based index
    }
  }
  
  // If we get here, no date range covers the current date
  return -1;
}
function setFirstCellValue(sheetName, value) {
  // Get the first cell in the sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var firstCell = sheet.getRange(1, 1);

  // Set the value and format of the cell
  firstCell.setValue(value);
  firstCell.setFontWeight("bold");
  firstCell.setNumberFormat("$#,##0.00");
}

