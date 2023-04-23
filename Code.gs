/*
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
let magic_sprintNames = "Sprint name";
let magic_scheduleSheetName = "Schedule";
let magic_summarySheetName = "Summary";

let index_col_outputInitiatives = 1;
let index_col_outputCurrentEstimates = 2;
let index_col_outputTotalInShedule = 3;
let index_col_outputDifferences = 4;
let index_col_outputSprints = 5;
let index_col_lookupDevNames = 1;
let index_col_lookupDevLocation = 2;
let index_col_lookupDevCost = 3;
let index_col_lookupInitiatives = 6;
let index_col_scheduleDevNames = 1;
let index_row_scheduleDevNamesBelow = 8;
let sheetName_lookup = "Lookups";
let sheetName_sprintHeader = "Sprint Header";
let index_row_scheduleSprintNameRow = 3;
let index_row_scheduleSprintStartDateRow = 4;
let index_row_scheduleSprintEndDateRow = 5;

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Do Le Calculationz")
    .addItem("calculate days per initiative", "calculateDaysPerInitiative")
    .addItem("calculate days per initiative per dev", "calculateEverything")
    .addToUi();
}

function getSprints(initiativesObj) {
  let scheduleSheet = SpreadsheetApp.getActive().getSheetByName(sheetName_sprintHeader);
  let scheduleRange = scheduleSheet.getDataRange();
  let lastColumn = scheduleRange.getLastColumn();
  let sprintNames = Array();
  let sprintNameIndexes = Array();
  let sprintDates = new Array();
  
  for (let c = 1; c < lastColumn - 1; c++) {
    let cell = scheduleRange.getCell(index_row_scheduleSprintNameRow, c);
    let sprintName = cell.getValue();

    if (sprintName != "" && sprintName != magic_sprintNames) {

      sprintNames.push(sprintName);
      sprintNameIndexes[c] = sprintName;

      let startDateCell = scheduleRange.getCell(index_row_scheduleSprintStartDateRow, c);
      let startDate = startDateCell.getValue();
      let endDateCell = scheduleRange.getCell(index_row_scheduleSprintEndDateRow, c);
      let endDate = endDateCell.getValue();

      let sprintNameAndDate = {
        Name: sprintName,
        Start: startDate,
        End: endDate,
        Cost: 0,
        Column: c,
        Days: 0,
        Initiatives: initialiseColourToCosts(initiativesObj)
      };
      sprintDates.push(sprintNameAndDate);
    }
  }

  let sprintObj = {};
  sprintObj.Names = sprintNames;
  sprintObj.NameIndexes = sprintNameIndexes;
  sprintObj = sprintDates;
  console.log("getSprints: sprintObj after updates", JSON.stringify(sprintObj));
  return sprintDates;
  return sprintObj;
}

function initialiseColourToCosts(initiativesObj) {
  let colourCosts = [];
  for (let i = 0; i < Object.keys(initiativesObj).length; i++) {
    let colour = Object.keys(initiativesObj)[i];
    colourCosts.push({ BackgroundColour: colour, Cost: 0, Days: 0 });
  }  
  return colourCosts;
}

function calculateEverything() {
  calculateDaysPerInitiative(true);
}

function sumInitiativeByBackgroundColour(scheduleSheetName, backgroundColour) {
  const refSheet = SpreadsheetApp.getActive().getSheetByName(scheduleSheetName);

  const startRow = index_row_scheduleDevNamesBelow + 1; // configurable start row
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
  // const initiatives = [];
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
  let initiatives = {};

  if (initiativeColumn > 0 && estimatesColumn > 0 && bgColumn > 0) {
    for (let ro = 1; ro < lookupsValues.length; ro++) {

      let initiativeName = lookupsValues[ro][initiativeColumn - 1];
      let initiativeCurrentEstimate = lookupsValues[ro][estimatesColumn - 1];
      let initiativeBgColour = lookupsValues[ro][bgColumn - 1];

      //initiatives.push(initiativeName);
      //estimates.push(initiativeCurrentEstimate);
      //bgColours.push(initiativeBgColour);
      if (initiativeName !== "")
      {
        let initiative = {
          Name: initiativeName,
          CurrentEstimate: initiativeCurrentEstimate,
          BackgroundColour: initiativeBgColour,
          TotalCost: 0,
          Sprints: [],
          Days: 0
        };
        initiatives[initiativeBgColour] = initiative;

      }
    }
  }
  console.log(JSON.stringify(initiatives));
  return initiatives;
  /*
  return {
    initiatives,
    estimates,
    bgColours
  };
  */
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


function getDevs(initiativesObj) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_lookup);
  const lastRow = sheet.getLastRow();
  const devData = sheet.getRange(`A2:C${lastRow}`).getValues();

  const devs = devData.reduce((result, [name, location, cost]) => {
    if (name !== "") {
      result[name] = {
        Name: name,
        Location: location,
        Cost: cost,
        TotalCost: 0,
        Sprints: [],
        Initiatives: initialiseColourToCosts(initiativesObj)
      };
    }
    return result;
  }, {});

  console.log(JSON.stringify(devs));
  console.log(Object.keys(devs).length);
  return devs;
}


function processSchedule(scheduleSheetName, devObj, sprintObj, initiativeObj) {
  let scheduleSheet = SpreadsheetApp.getActive().getSheetByName(scheduleSheetName);
  let scheduleRange = scheduleSheet.getRange(index_row_scheduleDevNamesBelow, 1, Object.keys(devObj).length + 1, scheduleSheet.getDataRange().getLastColumn());

  let sumOfDaysByBackgroundColour = {};
  let sprintNamesByInitiativeColour = {};
  const smallestSprintColumn = sprintObj.reduce((min, sprint) => {
    return sprint.Column < min ? sprint.Column : min;
  }, sprintObj[0].Column);
  console.log("smallestSprintColumn:",smallestSprintColumn);

  const largestSprintColumn = sprintObj.reduce((max, sprint) => {
    return sprint.Column > max ? sprint.Column : max;
  }, sprintObj[0].Column);
  console.log("largestSprintColumn:",largestSprintColumn);

  let devNamesBelowCell = scheduleRange.getCell(1, index_col_scheduleDevNames);
  let devNamesBelowCellContents = devNamesBelowCell.getValue();
  if (devNamesBelowCellContents != magic_DevNamesBelow) {
    Browser.msgBox("Your devs should be in row " + index_row_scheduleDevNamesBelow + " onwards, with the heading in row " + index_row_scheduleDevNamesBelow + " being '" + magic_DevNamesBelow + "', but you have the value " + devNamesBelowCellContents + " in row " + index_row_scheduleDevNamesBelow + ".");
    return;
  }

  for (let r = 2; r <= Object.keys(devObj).length + 1; r++) {
    console.log("ProcessSchedule: r", r);
    let devNameCell = scheduleRange.getCell(r, index_col_scheduleDevNames);
    let devName = devNameCell.getValue();
    console.log("ProcessSchedule: devName", devName);
    let thisDev = devObj[devName];
    console.log("ProcessSchedule: thisDev", JSON.stringify(thisDev));

    for (let c = smallestSprintColumn; c <= largestSprintColumn; c++) {
      console.log("ProcessSchedule: r,c", r + ", " + c);

      var thisSprintObj = sprintObj.find(s => s.Column == c);

      if (thisSprintObj !== undefined) {
        console.log("ProcessSchedule: thisSprintObj");
        logSprintDetails(thisSprintObj);

        let cell = scheduleRange.getCell(r, c);
        let cellBackground = cell.getBackground();
        let days = cell.getValue();
        console.log("ProcessSchedule: days", days);
        console.log("ProcessSchedule: cellBackground", cellBackground);

        if (days != null && days !== undefined && days !== "" && !isNaN(days)) {

          var thisInitiativeObj = initiativeObj[cellBackground];
          console.log("ProcessSchedule: thisInitiativeObj", JSON.stringify(thisInitiativeObj));

          let thisDevCostPerDay = thisDev.Cost;
          console.log("ProcessSchedule: thisDevCostPerDay", thisDevCostPerDay);
          let thisSprintCost = thisDevCostPerDay * days;
          console.log("ProcessSchedule: thisSprintCost", thisSprintCost);

          //increase the initiative days and cost
          thisInitiativeObj.Days += days;
          thisInitiativeObj.TotalCost += thisSprintCost;
          if (!Object.keys(thisInitiativeObj.Sprints).includes(thisSprintObj.Name)) {
            thisInitiativeObj.Sprints.push(thisSprintObj.Name);
          }
          console.log("ProcessSchedule: thisInitiativeObj after updates", JSON.stringify(thisInitiativeObj));

          //increase the sprint days
          thisSprintObj.Days += days;
          thisSprintObj.Cost += thisSprintCost;
          thisSprintObj.Initiatives.find(i => i.BackgroundColour == cellBackground).Cost += thisSprintCost;
          thisSprintObj.Initiatives.find(i => i.BackgroundColour == cellBackground).Days += days;
          console.log("ProcessSchedule: thisSprintObj after updates", JSON.stringify(thisSprintObj));

          //increase the developer days
          thisDev.Days += days;
          thisDev.TotalCost += thisSprintCost;
          thisDev.Initiatives.find(i => i.BackgroundColour == cellBackground).Cost += thisSprintCost;
          thisDev.Initiatives.find(i => i.BackgroundColour == cellBackground).Days += days;          
          if (!Object.keys(thisDev.Sprints).includes(thisSprintObj.Name)) {
            thisDev.Sprints.push(thisSprintObj.Name);
          }
          console.log("ProcessSchedule: thisDev after updates", JSON.stringify(thisDev));
        }
      }
    }
  }
}

function parseSheetName(sheetName) {
  console.log(sheetName);
  let summarySheetName = null;
  let scheduleSheetName = null;

  if (sheetName.endsWith(magic_summarySheetName) || sheetName.startsWith(magic_scheduleSheetName)) {
    if (sheetName.endsWith(magic_summarySheetName)) {
      scheduleSheetName = sheetName.slice(0, sheetName.indexOf(magic_summarySheetName));
      summarySheetName = sheetName;
    } else {
      scheduleSheetName = sheetName;
      summarySheetName = scheduleSheetName + magic_summarySheetName;
    }
  } else {
    return null;
  }

  return {
    summarySheetName: summarySheetName,
    scheduleSheetName: scheduleSheetName
  };
}

function calculateDaysPerInitiative(doCosts = false, rebuildSheet = false) {

  let thisSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  let sheets = parseSheetName(thisSheetName);
  console.log(sheets);

  if (sheets == null) {
    Browser.msgBox("Error", "Your request could not be completed, run the command from a sheet with a name starting with '" + magic_scheduleSheetName + "' or ending with '" + magic_summarySheetName + "'.", Browser.Buttons.OK);
    return;
  }

  let summarySheetName = sheets.summarySheetName;
  let scheduleSheetName = sheets.scheduleSheetName;

  var initiativesObj = getInitiatives();
  let sprintObj = getSprints(initiativesObj);
  var devsObj = getDevs(initiativesObj);
  processSchedule(scheduleSheetName, devsObj, sprintObj, initiativesObj);

  console.log("calculateDaysPerInitiative: sprintObj after processSchedule", JSON.stringify(sprintObj));
  console.log("calculateDaysPerInitiative: initiativesObj after processSchedule", JSON.stringify(initiativesObj));
  console.log("calculateDaysPerInitiative: devsObj after processSchedule", JSON.stringify(devsObj));
  throw new Error();

  let devsInSchedule = SpreadsheetApp.getActive().getSheetByName(sheetName_lookup).getRange(2, index_col_lookupDevNames, devsObj.length, 3);
  let summarySheet = createOrDeleteOrUpdateSummarySheet(summarySheetName, initiativesObj, rebuildSheet);

  let scheduleSheet = SpreadsheetApp.getActive().getSheetByName(scheduleSheetName);
  let scheduleRange = scheduleSheet.getRange(index_row_scheduleDevNamesBelow, 1, allDevCount, scheduleSheet.getDataRange().getLastColumn());
  let lastScheduleColumn = scheduleRange.getLastColumn();
  let devNames = Array();
  let differences = Array();
  let scheduledDays = Array();
  let sumOfDaysByBackgroundColour = Array();
  let devNameIndexes = Array();
  let sprintNamesByInitiativeColour = Array();

  let sprintNameByColumn = sprintObj["sprintNameIndexes"];
  let sprintNamesAndDates = sprintObj;

  var devCount = 0;
  for (let r = 1; r <= allDevCount; r++) {

    for (let c = 1; c < lastScheduleColumn; c++) {

      let thisSprintName = sprintNameByColumn[c];

      let cell = scheduleRange.getCell(r, c);
      let cellBackground = cell.getBackground();
      let days = cell.getValue();
      if (days != null && days !== undefined && days !== "") {

        if (r == 1 && c == 1 && days != magic_DevNamesBelow) {
          Browser.msgBox("Your devs should be in row " + index_row_scheduleDevNamesBelow + " onwards, with the heading in row " + index_row_scheduleDevNamesBelow + " being '" + magic_DevNamesBelow + "'.");
        }

        if (c == 1) {
          let devName = days;
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

        if (days != "" && !isNaN(days)) {
          if (!Object.keys(sumOfDaysByBackgroundColour).some(key => key === cellBackground)) {
            sumOfDaysByBackgroundColour[cellBackground] = 0;
          }
          sumOfDaysByBackgroundColour[cellBackground] += days;

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
          assignedDev["Initiatives"][cellBackground] += days;
          assignedDev["Total"] += days;
        }
      }
    }
  }

  let concatSprintNamesPerInitiative = Array();

  var summarySheetLastRow = summarySheet.getLastRow();
  for (let ro = 2; ro <= summarySheetLastRow; ro++) {
    let cell = summarySheet.getRange(ro, 1);
    let backgroundColour = cell.getBackground();
    let sumOfDays = sumOfDaysByBackgroundColour[backgroundColour];
    scheduledDays.push(sumOfDays);

    let estimateCell = summarySheet.getRange(ro, index_col_outputCurrentEstimates);
    let estimate = estimateCell.getValue();
    let difference = ((sumOfDays - estimate) / sumOfDays);
    differences.push(difference);

    let sprintsForThisInitiative = sprintNamesByInitiativeColour[backgroundColour];
    if (sprintsForThisInitiative != null) {
      uniq = arrayToCommaDelimitedList(sprintsForThisInitiative);
      concatSprintNamesPerInitiative.push(uniq);
    }
  }

  summarySheet.getRange(2, index_col_outputTotalInShedule, scheduledDays.length, 1).setValues(getValuesAs2DArray(scheduledDays));
  summarySheet.getRange(2, index_col_outputDifferences, differences.length, 1).setValues(getValuesAs2DArray(differences));
  summarySheet.getRange(2, index_col_outputSprints, concatSprintNamesPerInitiative.length, 1).setValues(getValuesAs2DArray(concatSprintNamesPerInitiative));

  var summarySheetLastColumn = summarySheet.getLastColumn();

  summarySheetLastRow = summarySheet.getLastRow();

  // for each dev in the schedule
  for (var devRowIndex = 1; devRowIndex < allDevCount; devRowIndex++) {
    var devToInitiative = Array();
    var cell = devsInSchedule.getCell(devRowIndex, 1);
    var devName = cell.getValue();
    var costCell = devsInSchedule.getCell(devRowIndex, 3);
    var cost = costCell.getValue();

    if (devName !== "") {
      devToInitiative.push(devName);
      var devObjectIndex = devNameIndexes[devName];
      var devObject = devNames[devObjectIndex];
      if (devObject != null) {

        devObject.cost = cost;

        // for each initiative in the output summary
        for (var initiativeColumnIndex = 2; initiativeColumnIndex <= summarySheetLastRow; initiativeColumnIndex++) {
          var initiativeNameCell = summarySheet.getRange(initiativeColumnIndex, index_col_outputInitiatives);
          var initiativeName = initiativeNameCell.getValue();
          var backgroundColour = initiativeNameCell.getBackground();
          var sumOfDaysThisInitiative = 0;

          if (Object.keys(devObject).some(key => key === "Initiatives") &&
            Object.keys(devObject["Initiatives"]).some(key => key === backgroundColour)) {
            sumOfDaysThisInitiative = devObject["Initiatives"][backgroundColour];
          }
          devToInitiative.push(sumOfDaysThisInitiative);
        }

        devToInitiative.push(devObject["Total"]);
        devToInitiative.push(devObject["Total"] * cost);
        devToInitiative.push();
      }
      else {
        Browser.msgBox("There is no information for: " + devName);
      }
    }
    summarySheet.getRange(1, summarySheetLastColumn + 1 + devRowIndex, devToInitiative.length, 1).setValues(getValuesAs2DArray(devToInitiative));
    console.log(JSON.stringify(devToInitiative));
  }

  if (doCosts == true) {

    logSprintsDetails(sprintNamesAndDates);
    updateSprintCosts(allDevCount, devNames, lastScheduleColumn, scheduleRange, sprintNamesAndDates);
    logSprintsDetails(sprintNamesAndDates);

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
      let cell = summarySheet.getRange(ro, index_col_outputInitiatives);
      let backgroundColour = cell.getBackground();

      // foreach dev in the lookup table
      for (let co = 1; co < devsInSchedule.getLastRow(); co++) {
        let cell = devsInSchedule.getCell(co, index_col_lookupDevNames);
        let devName = cell.getValue();
        let locationCell = devsInSchedule.getCell(co, index_col_lookupDevLocation);
        let location = locationCell.getValue();
        let costCell = devsInSchedule.getCell(co, index_col_lookupDevCost);
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

    // put the cost per initiative per sprint
    summarySheetLastRow = summarySheet.getLastRow();
    summarySheetLastColumn = summarySheet.getLastColumn() + 2;

    for (let sprintCount = 0; sprintCount < sprintNamesAndDates.length; sprintCount++) {
      let thisSprintInitiativeCosts = Array();
      for (let initiativeRow = 0; initiativeRow < initiativesObj.initiatives.length; initiativeRow++) {

        let cellRow = initiativeRow + 1;
        let thisSprint = sprintNamesAndDates[sprintCount];
        let cellColumn = thisSprint.sprintColumn;

        let cell = scheduleSheet.getRange(cellRow, cellColumn);
        let initiativeName = initiativesObj.initiatives[initiativeRow];
        if (sprintCount == 0) {
          if (initiativeRow == 0) {
            thisSprintInitiativeCosts.push(magic_Initiatives);
          }
          thisSprintInitiativeCosts.push(initiativeName);
        } else {
          let initiativeColour = cell.getBackground();

          if (initiativeRow == 0) {
            thisSprintInitiativeCosts.push(thisSprint.sprintName);
          }
          else {
            let costOfThisInitiativeThisSprint = thisSprint.Initiatives.find(c => c.BackgroundColour == initiativeColour);
            if (costOfThisInitiativeThisSprint !== undefined) {
              thisSprintInitiativeCosts.push(costOfThisInitiativeThisSprint.cost);
            } else {
              thisSprintInitiativeCosts.push(0);
            }
          }
        }
      }

      summarySheet.getRange(1, summarySheetLastColumn, thisSprintInitiativeCosts.length, 1).setValues(getValuesAs2DArray(thisSprintInitiativeCosts));
      summarySheetLastColumn++;
      console.log(JSON.stringify(thisSprintInitiativeCosts));
    }
  }

  summarySheet.autoResizeColumns(1, summarySheet.getLastColumn());
  var columnIndexes = [index_col_outputCurrentEstimates, index_col_outputTotalInShedule, index_col_outputDifferences];
  applyFormattingToSummarySheet(summarySheetName, initiativesObj, columnIndexes);

  let targetRow = allDevCount + index_row_scheduleDevNamesBelow + 3;
  var sourceRange = summarySheet.getRange('A1:D' + initiativesObj.initiatives.length);
  let bottomRowToUse = targetRow + 4;
  transposeDataWithFormat(scheduleSheetName, targetRow, sourceRange);

  if (doCosts) {
    sourceRange = getLastThreeColumnsRange(summarySheetName, initiativesObj.initiatives.length + 2); // also get the totals
    transposeDataWithFormat(scheduleSheetName, bottomRowToUse, sourceRange);
    bottomRowToUse += 3;

  }
  scheduleSheet.getRange(bottomRowToUse, 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).clearContent();
  scheduleSheet.getRange(scheduleSheet.getDataRange().getLastRow(), 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).setFontWeight("bold");
}

function logSprintDetails(columnSprint) {
  console.log(`Object: ${JSON.stringify(columnSprint)}`);
  console.log(`Sprint Name: ${columnSprint.Name}`);
  console.log(`Start Date: ${columnSprint.Start}`);
  console.log(`End Date: ${columnSprint.End}`);
  console.log(`Column: ${columnSprint.Column}`);
  console.log(`Days: ${columnSprint.Days}`);
  console.log(JSON.stringify(columnSprint.Initiatives));
}

function logSprintsDetails(sprintNamesAndDates) {
  console.log("sprintNamesAndDates:", JSON.stringify(sprintNamesAndDates));
  for (let i = 0; i < sprintNamesAndDates.length; i++) {
    let columnSprint = sprintNamesAndDates[i];
    logSprintDetails(columnSprint);
  }
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
  summarySheet.getRange(1, index_col_outputInitiatives, 1, 1).setValue(magic_Initiatives);
  summarySheet.getRange(1, index_col_outputCurrentEstimates, 1, 1).setValue(magic_currentEstimates);
  summarySheet.getRange(1, index_col_outputTotalInShedule, 1, 1).setValue(magic_totalInSchedule);
  summarySheet.getRange(1, index_col_outputDifferences, 1, 1).setValue(magic_Differences);
  summarySheet.getRange(1, index_col_outputSprints, 1, 1).setValue(magic_Sprints);
  summarySheet.getRange(2, index_col_outputInitiatives, initiatives.length, 1).setValues(getValuesAs2DArray(initiatives));
  summarySheet.getRange(2, index_col_outputCurrentEstimates, estimates.length, 1).setValues(getValuesAs2DArray(estimates));
  var columnIndexes = [index_col_outputDifferences];
  applyFormattingToSummarySheet(summarySheetName, initiativesObj, columnIndexes);
  return summarySheet;
}

function updateSummarySheet(summarySheetName, initiativesObj) {
  const summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheetName);
  const currentEstimates = initiativesObj.map(initiative => initiative.CurrentEstimate).join(', ');
  const initiativeNames = initiativesObj.map(initiative => initiative.Name).join(', ');

  summarySheet.getRange(2, index_col_outputInitiatives, initiativesObj.length, 1).setValues(getValuesAs2DArray(initiativeNames));
  summarySheet.getRange(2, index_col_outputCurrentEstimates, initiativesObj.length, 1).setValues(getValuesAs2DArray(currentEstimates));
  return summarySheet;
}

function applyFormattingToSummarySheet(summarySheetName, initiativesObj, columnIndexes) {
  var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheetName);
  const bgColours = initiativesObj.map(initiative => initiative.BackgroundColour).join(', ');

  for (let bg = 0; bg < bgColours.length; bg++) {
    summarySheet.getRange(bg + 2, index_col_outputInitiatives, 1, 1).setBackground(bgColours[bg]);
    summarySheet.getRange(bg + 2, index_col_outputDifferences, 1, 1).setNumberFormat('0.00%');
  }

  for (let i = 0; i < columnIndexes.length; i++) {
    var rangeA1Notation = getColumnRange(summarySheetName, columnIndexes[i]);
    applyConditionalFormatting(index_col_outputDifferences, rangeA1Notation, summarySheet);
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

function createOrDeleteOrUpdateSummarySheet(summarySheetName, initiativesObj, rebuildSheet) {
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
  return summarySheet;
}

function updateSprintCosts(allDevCount, devNames, lastScheduleColumn, scheduleRange, sprintNamesAndDates) {
  for (let r = 1; r <= allDevCount; r++) {
    var thisDev = devNames[r];
    var devCost = thisDev.cost;
    console.log("thisDev.Name:", JSON.stringify(thisDev.Name));
    console.log("devCost:", JSON.stringify(devCost));
    if (devCost > 0) {
      for (let c = 2; c <= lastScheduleColumn; c++) {
        console.log("r, c:", r + ", " + c);
        let cell = scheduleRange.getCell(r, c);
        let days = cell.getValue();
        console.log("days: ", JSON.stringify(days));
        if (days > 0) {
          var costOfThisDevThisInitiativeThisSprint = days * devCost;
          var columnSprint = sprintNamesAndDates.find(sprint => sprint.Column == c);
          if (columnSprint !== undefined) {

            logSprintDetails(columnSprint);

            columnSprint.sprintCost += costOfThisDevThisInitiativeThisSprint;
            columnSprint.sprintDays += days;
            let x = columnSprint.Initiatives.find(c => c.colour == cell.getBackground())
            if (x) {
              x.Cost += costOfThisDevThisInitiativeThisSprint;

              console.log("updateSprintCosts: costOfThisDevThisInitiativeThisSprint:", JSON.stringify(costOfThisDevThisInitiativeThisSprint));
              console.log(JSON.stringify(x));
              logSprintDetails(columnSprint);
            }
          }
          else {
            console.log("couldnt find sprint matching column", c)
          }
        }
      }
    }
  }
}

