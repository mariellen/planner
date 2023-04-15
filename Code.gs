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
  const range = refSheet.getRange("C10:Z30");
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

function UpdateSumInitiative(scheduleSheetName, row, col, rowtoupdate, coltoupdate) {
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
    .setRanges([range])
    .build();

  // Rule 2: Value between 0% and 20% is highlighted with an orange background
  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(${columnLetter}:${columnLetter}>=0%,${columnLetter}:${columnLetter}<20%)`)
    .setBackground("#ff8c00")
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


function calculateDaysPerInitiative(doCosts = false) {

  //Browser.msgBox(1);

  let scheduleSheetName = SpreadsheetApp.getActiveSheet().getName();
  let summarySheetName = scheduleSheetName + "Summary";

  // Use regular expression to check if sheet name starts with "Schedule" and does not end with "Summary"
  var pattern = /^schedule(?!.*summary$)/i; // i flag for case-insensitive matching
  if (!scheduleSheetName.match(pattern)) {
    Browser.msgBox("Error", "Your request could not be completed, run the command from a sheet with a name starting with 'Schedule' and not a Summary sheet.", Browser.Buttons.OK);

    return;
  }

  let summarySheet = SpreadsheetApp.getActive().getSheetByName(summarySheetName);
  /*
  if (summarySheet != null)
  {
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(summarySheet);
    summarySheet = null;
  }
  */

  if (summarySheet == null) {
    summarySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    summarySheet.setName(summarySheetName);
    const responseObj = getInitiatives();
    const initiatives = responseObj.initiatives;
    const estimates = responseObj.estimates;
    const bgColours = responseObj.bgColours;
    summarySheet.getRange(1, index_outputInitiatives, 1, 1).setValue(magic_Initiatives);
    summarySheet.getRange(1, index_outputCurrentEstimates, 1, 1).setValue(magic_currentEstimates);
    summarySheet.getRange(1, index_outputTotalInShedule, 1, 1).setValue(magic_totalInSchedule);
    summarySheet.getRange(1, index_outputDifferences, 1, 1).setValue(magic_Differences);
    summarySheet.getRange(1, index_outputSprints, 1, 1).setValue(magic_Sprints);
    summarySheet.getRange(2, 1, initiatives.length, 1).setValues(getValuesAs2DArray(initiatives));
    summarySheet.getRange(2, 2, estimates.length, 1).setValues(getValuesAs2DArray(estimates));

    for (let bg = 0; bg < bgColours.length; bg++) {
      summarySheet.getRange(bg + 2, 1, 1, 1).setBackground(bgColours[bg]);
      summarySheet.getRange(bg + 2, 4, 1, 1).setNumberFormat('0.00%');
    }
    let differenceRange = summarySheet.getRange(1, index_outputDifferences, summarySheet.getLastRow(), 1);
    applyConditionalFormatting(index_outputDifferences, differenceRange, summarySheet);
  }
  else {
    summarySheet.getRange(1, 5, summarySheet.getLastRow(), summarySheet.getLastColumn()).clearContent();
  }

  //Browser.msgBox(4);
  let scheduleSheet = SpreadsheetApp.getActive().getSheetByName(scheduleSheetName);
  let scheduleRange = scheduleSheet.getDataRange();
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

  for (let r = 11; r < lastScheduleRow - 1; r++) {
    for (let c = 1; c < lastScheduleColumn - 1; c++) {

      let thisSprintName = sprintNameByColumn[c];
      let cell = scheduleRange.getCell(r, c);
      let cellBackground = cell.getBackground();
      let cellValue = cell.getValue();

      if (r == 11 && c == 1 && cellValue != magic_DevNamesBelow) {
        Browser.msgBox("Your devs should be in row 11 onwards, with the heading in row 11 being '" + magic_DevNamesBelow + "'.");
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

  let concatSprintNamesPerInitiative = Array();

  let summarySheetLastRow = summarySheet.getLastRow();
  for (let ro = 2; ro <= summarySheetLastRow; ro++) {
    let cell = summarySheet.getRange(ro, 1);
    let backgroundColour = cell.getBackground();
    let sumOfDays = sumOfDaysByBackgroundColour[backgroundColour];
    scheduledDays.push(sumOfDays);

    if (index_outputCurrentEstimates > 0 && index_outputDifferences > 0) {

      let estimateCell = summarySheet.getRange(ro, index_outputCurrentEstimates);
      let estimate = estimateCell.getValue();
      let difference = ((sumOfDays - estimate) / sumOfDays);
      differences.push(difference);
    }

    let sprintsForThisInitiative = sprintNamesByInitiativeColour[backgroundColour];
    if (sprintsForThisInitiative != null) {
      uniq = arrayToCommaDelimitedList(sprintsForThisInitiative);
    }
    concatSprintNamesPerInitiative.push(uniq);
  }

  summarySheet.getRange(2, index_outputTotalInShedule, scheduledDays.length, 1).setValues(getValuesAs2DArray(scheduledDays));
  summarySheet.getRange(2, index_outputDifferences, differences.length, 1).setValues(getValuesAs2DArray(differences));
  summarySheet.getRange(2, index_outputSprints, concatSprintNamesPerInitiative.length, 1).setValues(getValuesAs2DArray(concatSprintNamesPerInitiative));

  let summarySheetLastColumn = summarySheet.getLastColumn();

  if (doCosts == true) {
    let allDevs = SpreadsheetApp.getActive().getSheetByName(sheetName_lookup).getRange(2, index_lookupDevNames, 20, 3);

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
    let rodevToInitiative = Array();
    let externaldevToInitiative = Array();
    let externalcostToInitiative = Array();
    let totalCost = 0;
    rodevToInitiative.push("ReachOut Days");
    externaldevToInitiative.push("External Days");
    externalcostToInitiative.push("External Cost");

    // foreach initiative in output sheet
    for (let ro = 2; ro <= summarySheetLastRow; ro++) {
      let reachOutDaysThisInitiative = 0;
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
        if (location == "ReachOut") {
          reachOutDaysThisInitiative += sumOfDaysThisInitiative;
        }
        else {
          externalDaysThisInitiative += sumOfDaysThisInitiative;
          externalCostThisInitiative += (sumOfDaysThisInitiative * cost);
        }
      }
      rodevToInitiative.push(reachOutDaysThisInitiative);
      externaldevToInitiative.push(externalDaysThisInitiative);
      externalcostToInitiative.push(externalCostThisInitiative);
      totalCost += externalCostThisInitiative;
    }
    externalcostToInitiative.push(totalCost);
    summarySheet.getRange(1, SpreadsheetApp.getActive().getSheetByName("Summary").getLastColumn() + 2, rodevToInitiative.length, 1).setValues(getValuesAs2DArray(rodevToInitiative));
    summarySheet.getRange(1, SpreadsheetApp.getActive().getSheetByName("Summary").getLastColumn() + 1, externaldevToInitiative.length, 1).setValues(getValuesAs2DArray(externaldevToInitiative));
    summarySheet.getRange(1, SpreadsheetApp.getActive().getSheetByName("Summary").getLastColumn() + 1, externalcostToInitiative.length, 1).setValues(getValuesAs2DArray(externalcostToInitiative));
  }

  summarySheet.autoResizeColumns(1, summarySheet.getLastColumn());

  transposeDataWithFormatAndFormatting(summarySheetName, scheduleSheetName, scheduleSheet.getLastRow() + 3);

}

function transposeDataWithFormatAndFormatting(sourceSheetName, targetSheetName, targetRow) {
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceSheetName); // Get the source sheet by name
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName); // Get the target sheet by name

  var sourceRange = sourceSheet.getRange('A:F');
  var sourceData = sourceRange.getValues(); // Get data from range A:F in the source sheet
  var sourceFormat = sourceRange.getNumberFormats(); // Get number formats from range A:F in the source sheet
  var sourceStyle = sourceRange.getRichTextValues(); // Get cell style formatting from range A:F in the source sheet
  var sourceConditionalFormat = sourceRange.getConditionalFormatRules(); // Get conditional formatting rules from range A:F in the source sheet

  var transposedData = transposeArray(sourceData); // Transpose the data
  var transposedFormat = transposeArray(sourceFormat); // Transpose the number formats
  var transposedStyle = transposeArray(sourceStyle); // Transpose the cell style formatting
  var transposedConditionalFormat = transposeArray(sourceConditionalFormat); // Transpose the conditional formatting rules

  var targetRange = targetSheet.getRange(targetRow, 1, transposedData.length, transposedData[0].length);
  targetRange.clearContent(); // Clear the contents of the target range
  targetRange.setValues(transposedData); // Insert the transposed data into the target sheet
  targetRange.setNumberFormats(transposedFormat); // Set the number formats in the target sheet

  for (var i = 0; i < transposedStyle.length; i++) {
    for (var j = 0; j < transposedStyle[i].length; j++) {
      targetRange.getCell(i+1, j+1).setRichTextValue(transposedStyle[i][j]); // Set cell style formatting in the target sheet
    }
  }

  targetSheet.setConditionalFormatRules(transposedConditionalFormat); // Set conditional formatting rules in the target sheet
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


