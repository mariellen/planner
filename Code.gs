let magic_HeadingTotalInSchedule = "Total in Schedule";
let magic_HeadingTotalInScheduleFromCurrentSprint = "From";
let magic_HeadingEstimateRemaining = "Remaining";
let magic_HeadingCurrentEstimates = "Current Estimate";
let magic_HeadingDifferences = "Difference";
let magic_HeadingSprints = "Sprints";
let magic_HeadingInitiativeCompletedInSprints = "Completed By";
let magic_HeadingInitiatives = "Initiative";
let magic_Background = "Background";
let magic_DevNamesBelow = "Dev Names Below";
let magic_sprintNames = "Sprint name";
let magic_scheduleSheetName = "Schedule";
let magic_summarySheetName = "Summary";
let magic_namedRangeIdentifier = "_Range";
let magic_InternalDays = "Internal Days";
let magic_ExternalDays = "External Days";
let magic_ExternalCost = "External Cost";
let magic_ExternalCostFromCurrentSprint = "External Cost from Current Sprint";
let magic_LocationInternal = "Internal";
let magic_RangeTypeEstimates = "Estimates";
let magic_RangeTypeSprintCosts = "SprintCosts";
let magic_RangeTypeInitiativeCosts = "InitiativeCosts";
let index_col_outputInitiatives = 1;
let index_col_outputCurrentEstimates = index_col_outputInitiatives + 1;
let index_col_outputTotalInShedule = index_col_outputCurrentEstimates + 1;
let index_col_outputDifferences = index_col_outputTotalInShedule + 1;
let index_col_outputEstimateRemaining = index_col_outputDifferences + 1;
let index_col_outputTotalInSheduleFromCurrentSprint = index_col_outputEstimateRemaining + 1;
let index_col_outputDifferencesFromCurrentSprint = index_col_outputTotalInSheduleFromCurrentSprint + 1;
let index_col_outputSprints = index_col_outputDifferencesFromCurrentSprint + 1;
let index_col_outputInitiativeCompleteSprint = index_col_outputSprints + 1;
let index_col_lookupDevNames = 1;
let index_col_lookupDevLocation = 2;
let index_col_lookupDevCost = 3;
let index_col_lookupInitiatives = 6;
let index_col_scheduleDevNames = 1;
let index_row_scheduleDevNamesBelow = 8;
let sheetName_lookup = "Lookups"; //only one
let sheetName_sprintHeader = "Sprint Header"; //only one
let sheetName_ReviewAll = "Review All";
let index_row_scheduleSprintNameRow = 3;
let index_row_scheduleSprintStartDateRow = 4;
let index_row_scheduleSprintEndDateRow = 5;

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Do Le Calculationz")
    .addItem("calculate days per initiative on Active Sheet", "calculateDaysPerInitiative")
    .addItem("calculate days per initiative per dev on Active Sheet", "calculateEverything")
    .addItem("calculate all estimates", "calculateAllEstimates")
    .addItem("calculate all costs", "calculateAllCosts")
    .addItem("collate review", "collateReview")
    .addToUi();
}
function getAllNamedRanges() {
  var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  let estimateRanges = Array();
  let initiativeCostRanges = Array();
  let sprintCostRanges = Array();
  let costsBySchedule = Array();

  for (var i = 0; i < namedRanges.length; i++) {
    let rangeName = namedRanges[i].getName();
    var scheduleSheetName = rangeName.slice(0, rangeName.indexOf(magic_namedRangeIdentifier));
    if (!Object.keys(costsBySchedule).includes(scheduleSheetName) && SpreadsheetApp.getActiveSpreadsheet().getSheetByName(scheduleSheetName + magic_summarySheetName) != null) {
      costsBySchedule[scheduleSheetName] = {
        Estimates: "",
        InitiativeCosts: "",
        SprintCosts: ""
      }
    }
    if (rangeName.endsWith(magic_RangeTypeEstimates)) {
      estimateRanges.push(rangeName);
      costsBySchedule[scheduleSheetName].Estimates = rangeName;
    }
    if (rangeName.endsWith(magic_RangeTypeInitiativeCosts)) {
      initiativeCostRanges.push(rangeName);
      costsBySchedule[scheduleSheetName].InitiativeCosts = rangeName;
    }
    if (rangeName.endsWith(magic_RangeTypeSprintCosts)) {
      sprintCostRanges.push(rangeName);
      costsBySchedule[scheduleSheetName].SprintCosts = rangeName;
    }
  }
  return [estimateRanges, initiativeCostRanges, sprintCostRanges, costsBySchedule];
}
function excludeEmptyColumns(range, headingPrefix) {

  var values = range.getValues(); // get the values in the range
  var numRows = values.length;
  var numCols = values[0].length;
  var newValues = [];

  for (var j = 0; j < numCols; j++) {
    var hasNonEmptyRows = false;

    for (var i = 0; i < numRows; i++) {
      console.log("j", j);
      console.log("i", i);
      console.log("hasNonEmptyRows", values[i][j]);

      if (values[i][j] !== "") {
        hasNonEmptyRows = true;
        console.log("hasNonEmptyRows", values[i][j]);
        break;
      }
    }

    if (hasNonEmptyRows) {
      var column = [];

      for (var i = 0; i < numRows; i++) {
        var val = values[i][j];
        column.push(val);
      }

      newValues.push(column);
    }
  }
  return newValues;
}


function copyNamedRangesToSheet(rangeNames) {
  var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_ReviewAll);
  const sortedRanges = Object.keys(rangeNames).sort();
  let rowToUse = 1;
  let columnToUse = 1;
  for (var i = 0; i < sortedRanges.length; i++) {
    var scheduleSheetName = sortedRanges[i];
    destinationSheet.getRange(rowToUse, columnToUse, 1, 1).setValue(scheduleSheetName);
    for (var r = 0; r < 3; r++) {
      var rangeName = Object.values(rangeNames[scheduleSheetName])[r];
      if (!rangeName.endsWith(magic_RangeTypeSprintCosts)) {
        var namedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName);
        var rangeWithoutEmptyColumns = excludeEmptyColumns(namedRange, scheduleSheetName);
        for (var rr = 0; rr < rangeWithoutEmptyColumns[0].length; rr++) {
          for (var cc = 0; cc < rangeWithoutEmptyColumns.length; cc++) {
            if (cc == 0) { // if current row is first row, add headingPrefix
              rangeWithoutEmptyColumns[cc][rr] = scheduleSheetName + " " + rangeWithoutEmptyColumns[cc][rr];
            }
          }
        }
        var numRows = rangeWithoutEmptyColumns.length;
        if (rangeWithoutEmptyColumns[0] != null) {
          var numCols = rangeWithoutEmptyColumns[0].length;
          destinationSheet.getRange(rowToUse, columnToUse, numRows, numCols).setValues(rangeWithoutEmptyColumns);
          columnToUse += numCols;
        }
      }
    }
  }
}

function copyNamedRangesToSheet2(costsBySchedule) {
  var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_ReviewAll);

  for (var i = 0; i < Object.keys(costsBySchedule).length; i++) {
    var scheduleName = Object.keys(costsBySchedule)[i];
    var estimateRangeName = costsBySchedule[scheduleName].Estimates;
    var initiativeCostsRangeName = costsBySchedule[scheduleName].InitiativeCosts;
    var sprintCostsRangeName = costsBySchedule[scheduleName].SprintCosts;

    var estimateRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(estimateRangeName);
    var initiativeCostsRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(initiativeCostsRangeName);
    var sprintCostsRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(sprintCostsRangeName);
    var estimateRangeValues = estimateRange.getValues();
    var initiativeCostsRangeValues = initiativeCostsRange.getValues();
    var sprintCostsRangeValues = sprintCostsRange.getValues();

    var numRows = estimateRangeValues.length;
    var numCols = estimateRangeValues[0].length;
    let rowToUse = destinationSheet.getLastRow() + 1;
    destinationSheet.getRange(rowToUse, 1, 1, 1).setValue(scheduleName + " " + magic_RangeTypeEstimates);
    destinationSheet.getRange(rowToUse, 2, numRows, numCols).setValues(estimateRangeValues);

    numRows = initiativeCostsRangeValues.length;
    numCols = initiativeCostsRangeValues[0].length;
    rowToUse = destinationSheet.getLastRow() + 1;
    destinationSheet.getRange(rowToUse, 1, 1, 1).setValue(scheduleName + " " + magic_RangeTypeInitiativeCosts);
    destinationSheet.getRange(rowToUse, 2, numRows, numCols).setValues(initiativeCostsRangeValues);

    numRows = sprintCostsRangeValues.length;
    numCols = sprintCostsRangeValues[0].length;
    rowToUse = destinationSheet.getLastRow() + 1;
    destinationSheet.getRange(rowToUse, 1, 1, 1).setValue(scheduleName + " " + magic_RangeTypeSprintCosts);
    destinationSheet.getRange(rowToUse, 2, numRows, numCols).setValues(sprintCostsRangeValues);
  }
}
function collateReview() {
  var reviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName_ReviewAll);
  reviewSheet.getDataRange().clearContent();
  reviewSheet.getDataRange().clearFormat();
  let [estimateRanges, initiativeCostRanges, sprintCostRanges, costsBySchedule] = getAllNamedRanges();
  copyNamedRangesToSheet(costsBySchedule);
}


function getSprints(initiativesObj) {
  let scheduleSheet = SpreadsheetApp.getActive().getSheetByName(sheetName_sprintHeader);
  let scheduleRange = scheduleSheet.getDataRange();
  let lastColumn = scheduleRange.getLastColumn();
  let sprintDates = Array();

  for (let c = 1; c < lastColumn - 1; c++) {
    let cell = scheduleRange.getCell(index_row_scheduleSprintNameRow, c);
    let sprintName = cell.getValue();

    if (sprintName != "" && sprintName != magic_sprintNames) {

      let startDateCell = scheduleRange.getCell(index_row_scheduleSprintStartDateRow, c);
      let startDate = startDateCell.getValue();
      let endDateCell = scheduleRange.getCell(index_row_scheduleSprintEndDateRow, c);
      let endDate = endDateCell.getValue();

      let sprintNameAndDate = {
        Name: sprintName,
        Start: startDate,
        End: endDate,
        Cost: 0,
        Columns: [],
        Index: c,
        Days: 0,
        Initiatives: initialiseColourToCosts(initiativesObj)
      };
      sprintDates.push(sprintNameAndDate);
    }
  }
  return sprintDates;
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
function getValuesAs2DArray(things) {
  let thingArray = Array();
  for (r = 0; r < things.length; r++) {
    thingArray[r] = Array();
    thingArray[r][0] = things[r];
  }
  return thingArray;
}

function getInitiatives() {

  let initiativeColumn = 0;
  let bgColumn = 0;
  let estimatesColumn = 0;
  let remainingColumn = 0;
  const lookupsSheet = SpreadsheetApp.getActive().getSheetByName(sheetName_lookup);
  const lookupsDataRange = lookupsSheet.getDataRange();
  const lookupsValues = lookupsDataRange.getValues();
  const columnValues = lookupsValues[0];
  columnValues.forEach((value, index) => {
    if (value === magic_HeadingInitiatives) {
      initiativeColumn = index + 1;
    }
    if (value === magic_HeadingEstimateRemaining) {
      remainingColumn = index + 1;
    }
    if (value === magic_HeadingCurrentEstimates) {
      estimatesColumn = index + 1;
    }
    if (value === magic_Background) {
      bgColumn = index + 1;
    }
  });
  let initiatives = [];

  if (initiativeColumn > 0 && estimatesColumn > 0 && bgColumn > 0) {
    for (let ro = 1; ro < lookupsValues.length; ro++) {

      let initiativeName = lookupsValues[ro][initiativeColumn - 1];
      let initiativeRemaining = lookupsValues[ro][remainingColumn - 1];
      let initiativeCurrentEstimate = lookupsValues[ro][estimatesColumn - 1];
      let initiativeBgColour = lookupsValues[ro][bgColumn - 1];

      if (initiativeName !== "") {
        let initiative = {
          Name: initiativeName,
          CurrentEstimate: initiativeCurrentEstimate,
          Remaining: initiativeRemaining,
          BackgroundColour: initiativeBgColour,
          TotalCost: 0,
          CostFromCurrentSprint: 0,
          Sprints: [],
          Days: 0,
          InternalDays: 0,
          ExternalDays: 0,
          ExternalCost: 0
        };
        initiatives[initiativeBgColour] = initiative;

      }
    }
  }
  console.log(JSON.stringify(initiatives));
  return initiatives;
}

function getBackgroundColourForInitiative(row, col) {
  let backgroundColour = SpreadsheetApp.getActiveSheet().getRange(row, col).getBackground();
  return backgroundColour;
}

function applyConditionalFormatting(columnIndex, range, sheet) {
  const columnLetter = getColumnLetter(columnIndex);

  // Rule 1: Value less than 20% or greater than 40% is highlighted with a dark red background
  const rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=OR(${columnLetter}:${columnLetter}<20%,${columnLetter}:${columnLetter}>30%)`)
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
    .whenFormulaSatisfied(`=AND(${columnLetter}:${columnLetter}>=20%,${columnLetter}:${columnLetter}<=30%)`)
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
        CostFromCurrentSprint: 0,
        Sprints: [],
        Initiatives: initialiseColourToCosts(initiativesObj),
        Days: 0,
        RowInSchedule: 0
      };
    }
    return result;
  }, {});

  console.log(JSON.stringify(devs));
  console.log(Object.keys(devs).length);
  return devs;
}

function setColumnsForSprints(scheduleSheetName, sprintObj) {
  let scheduleSheet = SpreadsheetApp.getActive().getSheetByName(scheduleSheetName);
  const headingsRow = scheduleSheet.getRange(index_row_scheduleSprintNameRow, 1, 1, scheduleSheet.getLastColumn()).getValues()[0];

  headingsRow.forEach((heading, columnIndex) => {
    const range = scheduleSheet.getRange(index_row_scheduleSprintNameRow, columnIndex + 1);
    const sprint = sprintObj.find(sprint => sprint.Name === heading);
    if (sprint) {
      if (!sprint.Columns) {
        sprint.Columns = [];
      }
      if (range.isPartOfMerge()) {
        const mergedRanges = range.getMergedRanges();
        // If merged, add each column in merged range to sprint.Columns
        mergedRanges.forEach(mergedRange => {
          const startColumnIndex = mergedRange.getColumn();
          const endColumnIndex = mergedRange.getLastColumn();

          for (let i = startColumnIndex; i <= endColumnIndex; i++) {
            sprint.Columns.push(i);
          }
        });
      }
      else {
        sprint.Columns.push(columnIndex + 1);
      }
    }
  });
  console.log("setColumnsForSprints: sprintObj", JSON.stringify(sprintObj));
  return sprintObj;
}
function getSmallestLargestSprintColumns(sprintObj) {
  let smallestColumn = Infinity;
  let largestColumn = 0;
  for (let i = 0; i < sprintObj.length; i++) {
    let sprint = sprintObj[i];
    for (let j = 0; j < sprint.Columns.length; j++) {
      let column = sprint.Columns[j];
      if (column < smallestColumn) {
        smallestColumn = column;
      }
      if (column > largestColumn) {
        largestColumn = column;
      }
    }
  }
  return [smallestColumn, largestColumn];
}

function processSchedule(scheduleSheetName, devObj, sprintObj, initiativeObj) {
  let scheduleSheet = SpreadsheetApp.getActive().getSheetByName(scheduleSheetName);
  let scheduleSheetLastColumn = scheduleSheet.getDataRange().getLastColumn();
  let scheduleRange = scheduleSheet.getRange(index_row_scheduleDevNamesBelow, 1, Object.keys(devObj).length + 1, scheduleSheetLastColumn);

  var currentSprintObj = getCurrentSprintObjByDate(sprintObj);

  sprintObj = setColumnsForSprints(scheduleSheetName, sprintObj);
  let [smallestSprintColumn, largestSprintColumn] = getSmallestLargestSprintColumns(sprintObj);

  let devNamesBelowCell = scheduleRange.getCell(1, index_col_scheduleDevNames);
  let devNamesBelowCellContents = devNamesBelowCell.getValue();
  if (devNamesBelowCellContents != magic_DevNamesBelow) {
    Browser.msgBox("Your devs should be in row " + index_row_scheduleDevNamesBelow + " onwards, with the heading in row " + index_row_scheduleDevNamesBelow + " being '" + magic_DevNamesBelow + "', but you have the value " + devNamesBelowCellContents + " in row " + index_row_scheduleDevNamesBelow + ".");
    return;
  }

  // loping through each dev
  for (let r = 2; r <= Object.keys(devObj).length + 1; r++) {

    console.log("MARIELLEN TODO", "ONLY BOTHER WHEN THE SUM of the ROW is > 0");
    console.log("ProcessSchedule: r", r);
    let devNameCell = scheduleRange.getCell(r, index_col_scheduleDevNames);
    let devName = devNameCell.getValue();
    console.log("ProcessSchedule: devName", devName);
    let thisDev = devObj[devName];
    thisDev.RowInScheduleRange = r;
    console.log("ProcessSchedule: thisDev", JSON.stringify(thisDev));

    // looping through each sprint
    for (let c = smallestSprintColumn; c <= largestSprintColumn; c++) {
      console.log("ProcessSchedule: r,c", r + ", " + c);

      let scheduleSprintName = scheduleRange.getCell(index_row_scheduleSprintNameRow, c);
      var thisSprintObj = sprintObj.find(s => s.Columns.find(col => col == c) || s.Name == scheduleSprintName);

      if (thisSprintObj !== undefined) {
        console.log("ProcessSchedule: thisSprintObj before updates");
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

          if (!(new Date(thisSprintObj.End).getTime() <= new Date(currentSprintObj.Start).getTime())) {
            thisInitiativeObj.CostFromCurrentSprint += thisSprintCost;
          }
          if (!Object.values(thisInitiativeObj.Sprints).includes(thisSprintObj.Name)) {
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
          if (!(new Date(thisSprintObj.End).getTime() <= new Date(thisSprintObj.Start).getTime())) {
            thisDev.CostFromCurrentSprint += thisSprintCost;
          }
          thisDev.Initiatives.find(i => i.BackgroundColour == cellBackground).Cost += thisSprintCost;
          thisDev.Initiatives.find(i => i.BackgroundColour == cellBackground).Days += days;
          if (!Object.keys(thisDev.Sprints).includes(thisSprintObj.Name)) {
            thisDev.Sprints[thisSprintObj.Name] = {
              Name: thisSprintObj.Name,
              Days: days
            }
          }
          else {
            thisDev.Sprints[thisSprintObj.Name].Days += days;
          }

          if (thisDev.Location == magic_LocationInternal) {
            thisInitiativeObj.InternalDays += days;
          }
          else {
            thisInitiativeObj.ExternalDays += days;
            thisInitiativeObj.ExternalCost += thisSprintCost;
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

function populateSummaryWithEstimates(summarySheetName, initiativesObj, sprintObj, scheduleSheetName) {
  let summarySheet = SpreadsheetApp.getActive().getSheetByName(summarySheetName);

  let concatSprintNamesPerInitiative = Array();
  let differences = Array();
  let differencesFromCurrentSprint = Array();
  let scheduledDays = Array();
  let totalCurrentEstimates = 0;
  let initiativeCompleteSprint = Array();
  let scheduledDaysFromCurrentSprint = Array();
  let remaining = Array();
  let totalRemaining = 0;

  var currentSprintCol = getCurrentDateRangeColumn(scheduleSheetName);
  var currentSprintObject = sprintObj.find(sp => sp.Columns.includes(currentSprintCol));
  var currentSprintStartDate = currentSprintObject.Start;

  for (let ro = 2; ro <= Object.keys(initiativesObj).length + 1; ro++) {
    let initiativeCell = summarySheet.getRange(ro, 1);
    let backgroundColour = initiativeCell.getBackground();
    let thisInitiative = initiativesObj[backgroundColour];

    if (thisInitiative != null) {

      var initiativeSprints = thisInitiative.Sprints;

      let sumOfDays = thisInitiative.Days;
      scheduledDays.push(sumOfDays);
      totalCurrentEstimates += thisInitiative.CurrentEstimate;

      let estimate = thisInitiative.CurrentEstimate;
      let difference = ((sumOfDays - estimate) / sumOfDays);
      differences.push(difference);

      let remainingNow = thisInitiative.Remaining;
      remaining.push(remainingNow);
      totalRemaining += remainingNow;



      var sprintThing = Object.values(initiativeSprints);
      var lastSprint = sprintThing[sprintThing.length - 1];
      concatSprintNamesPerInitiative.push(sprintThing.map(sprint => sprint).join(', '));
      initiativeCompleteSprint.push(lastSprint);

      let totalFromCurrentSprint = 0
      for (let s = 0; s < Object.keys(initiativeSprints).length; s++) {
        let thisKey = Object.keys(initiativeSprints)[s];
        let loopingSprintName = initiativeSprints[thisKey];
        let actualSprint = sprintObj.find(so => so.Name == loopingSprintName);

        if (!(new Date(actualSprint.End).getTime() <= new Date(currentSprintStartDate).getTime())) {
          var sprintInitiativeDeets = sprintObj.find(so => so.Name == loopingSprintName);
          var sprintInitiatives = sprintInitiativeDeets.Initiatives.find(si => si.BackgroundColour == backgroundColour);
          totalFromCurrentSprint += sprintInitiatives.Days;
        }
      }
      scheduledDaysFromCurrentSprint.push(totalFromCurrentSprint);

      let differenceFromCurrentSprint = ((totalFromCurrentSprint - remainingNow) / totalFromCurrentSprint);
      differencesFromCurrentSprint.push(differenceFromCurrentSprint);

    }

  }
  remaining.push(totalRemaining);

  if (scheduledDays.length === 0) {

    scheduledDays.push("-");

  }
  else {
    var totalDaysInSchedule = scheduledDays.reduce((accumulator, currentValue) => accumulator + currentValue);
    scheduledDays.push(totalDaysInSchedule);

  }

  if (scheduledDaysFromCurrentSprint.length === 0) {
    scheduledDaysFromCurrentSprint.push("-");
  }
  else {
    var totalDaysInScheduleFromCurrentSprint = scheduledDaysFromCurrentSprint.reduce((accumulator, currentValue) => accumulator + currentValue);
    scheduledDaysFromCurrentSprint.push(totalDaysInScheduleFromCurrentSprint);
  }

  if (concatSprintNamesPerInitiative.length != 0) {
    summarySheet.getRange(2, index_col_outputSprints, concatSprintNamesPerInitiative.length, 1).setValues(getValuesAs2DArray(concatSprintNamesPerInitiative));
  }

  summarySheet.getRange(2, index_col_outputTotalInShedule, scheduledDays.length, 1).setValues(getValuesAs2DArray(scheduledDays));
  summarySheet.getRange(2, index_col_outputTotalInSheduleFromCurrentSprint, scheduledDaysFromCurrentSprint.length, 1).setValues(getValuesAs2DArray(scheduledDaysFromCurrentSprint));
  if (differences.length != 0) {
    summarySheet.getRange(2, index_col_outputDifferences, differences.length, 1).setValues(getValuesAs2DArray(differences));
  }
  if (differencesFromCurrentSprint.length != 0) {
    summarySheet.getRange(2, index_col_outputDifferencesFromCurrentSprint, differences.length, 1).setValues(getValuesAs2DArray(differencesFromCurrentSprint));
  }

  if (initiativeCompleteSprint.length != 0) {
    summarySheet.getRange(2, index_col_outputInitiativeCompleteSprint, initiativeCompleteSprint.length, 1).setValues(getValuesAs2DArray(initiativeCompleteSprint));
  }
  if (remaining.length != 0) {
    summarySheet.getRange(2, index_col_outputEstimateRemaining, remaining.length, 1).setValues(getValuesAs2DArray(remaining));
  }

  summarySheet.getRange(Object.keys(initiativesObj).length + 2, index_col_outputCurrentEstimates, 1, 1).setValue(totalCurrentEstimates);
}

function populateDevDaysPerInitiative(summarySheetName, devObj, initiativesObj) {
  let summarySheet = SpreadsheetApp.getActive().getSheetByName(summarySheetName);
  var summarySheetLastColumn = summarySheet.getLastColumn();

  for (var devCounter = 0; devCounter < Object.keys(devObj).length; devCounter++) {
    var devToInitiative = Array();
    var devName = Object.keys(devObj)[devCounter];
    devToInitiative.push(devName);
    var thisDev = devObj[devName];

    for (var initiativeCounter = 0; initiativeCounter < Object.keys(initiativesObj).length; initiativeCounter++) {
      var thisInitiativeColour = Object.keys(initiativesObj)[initiativeCounter];
      var thisDevDaysOnInitiative = thisDev.Initiatives.find(initiative => initiative.BackgroundColour == thisInitiativeColour);
      devToInitiative.push(thisDevDaysOnInitiative.Days);
    }
    devToInitiative.push(thisDev.Days);
    devToInitiative.push(thisDev.TotalCost);
    devToInitiative.push(thisDev.CostFromCurrentSprint);
    devToInitiative.push();
    summarySheet.getRange(1, summarySheetLastColumn + 1 + devCounter, devToInitiative.length, 1).setValues(getValuesAs2DArray(devToInitiative));
    console.log(JSON.stringify(devToInitiative));
  }
}
function getInitiativeRange(summarySheetName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(summarySheetName);
  const numRows = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const firstRowValues = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const secondInitiativeIndex = firstRowValues.indexOf(magic_HeadingInitiatives, firstRowValues.indexOf(magic_HeadingInitiatives) + 1);
  const range = sheet.getRange(1, secondInitiativeIndex + 1, numRows, lastColumn - secondInitiativeIndex);
  return range;
}

function populateInternalAndExternalCosts(summarySheetName, scheduleSheetName, initiativesObj, devObj) {
  let summarySheet = SpreadsheetApp.getActive().getSheetByName(summarySheetName);
  var summarySheetLastColumn = summarySheet.getLastColumn();

  let internaldevToInitiative = Array();
  let externaldevToInitiative = Array();
  let externalcostToInitiative = Array();
  let externalcostToInitiativeFromCurrentSprint = Array();

  internaldevToInitiative.push(magic_InternalDays);
  externaldevToInitiative.push(magic_ExternalDays);
  externalcostToInitiative.push(magic_ExternalCost);
  externalcostToInitiativeFromCurrentSprint.push(magic_ExternalCostFromCurrentSprint);

  let totalExternalCost = 0;
  let totalExternalCostFromCurrentSprint = 0;
  let totalExternalDays = 0;
  let totalInternalDays = 0;

  for (var initiativeCounter = 0; initiativeCounter < Object.keys(initiativesObj).length; initiativeCounter++) {
    var backgroundColour = Object.keys(initiativesObj)[initiativeCounter];
    internaldevToInitiative.push(initiativesObj[backgroundColour].InternalDays);
    externaldevToInitiative.push(initiativesObj[backgroundColour].ExternalDays);
    externalcostToInitiative.push(initiativesObj[backgroundColour].ExternalCost);
    externalcostToInitiativeFromCurrentSprint.push(initiativesObj[backgroundColour].CostFromCurrentSprint);

    totalExternalCostFromCurrentSprint += initiativesObj[backgroundColour].CostFromCurrentSprint;
    totalExternalCost += initiativesObj[backgroundColour].ExternalCost;
    totalExternalDays += initiativesObj[backgroundColour].ExternalDays;
    totalInternalDays += initiativesObj[backgroundColour].InternalDays;
  }

  externalcostToInitiativeFromCurrentSprint.push(totalExternalCostFromCurrentSprint);
  externalcostToInitiative.push(totalExternalCost);
  externalcostToInitiative.push("<-- Total Cost Per Dev");
  externalcostToInitiative.push("<-- Cost Per Dev From Current Sprint");
  externaldevToInitiative.push(totalExternalDays);
  internaldevToInitiative.push(totalInternalDays);

  summarySheetLastColumn = summarySheet.getLastColumn();
  summarySheet.getRange(1, summarySheetLastColumn + 1, internaldevToInitiative.length, 1).setValues(getValuesAs2DArray(internaldevToInitiative));
  summarySheet.getRange(1, summarySheetLastColumn + 2, externaldevToInitiative.length, 1).setValues(getValuesAs2DArray(externaldevToInitiative));
  summarySheet.getRange(1, summarySheetLastColumn + 3, externalcostToInitiative.length, 1).setValues(getValuesAs2DArray(externalcostToInitiative));
  summarySheet.getRange(1, summarySheetLastColumn + 4, externalcostToInitiativeFromCurrentSprint.length, 1).setValues(getValuesAs2DArray(externalcostToInitiativeFromCurrentSprint));

  summarySheet.getRange(summarySheet.getLastRow(), 1, 1, summarySheet.getLastColumn()).setNumberFormat("$00.00");
  summarySheet.getRange(summarySheet.getLastRow(), 1, 1, summarySheet.getLastColumn()).setFontWeight("bold");
  summarySheet.getRange(1, summarySheet.getLastColumn() - 1, summarySheet.getLastRow(), 2).setNumberFormat("$00.00");
  summarySheet.getRange(1, summarySheet.getLastColumn() - 1, summarySheet.getLastRow(), 2).setFontWeight("bold");

  // UPDATE THE SCHEDULE SHEET

  setFirstCellValue(scheduleSheetName, totalExternalCost);
  let scheduleSheet = SpreadsheetApp.getActive().getSheetByName(scheduleSheetName);
  var bottomRowToUse = Object.keys(devObj).length + index_row_scheduleDevNamesBelow + 6;

  sourceRange = getLastThreeColumnsRange(summarySheetName, Object.keys(initiativesObj).length + 2); // also get the totals
  transposeDataWithFormat(scheduleSheetName, bottomRowToUse, sourceRange, magic_RangeTypeInitiativeCosts);

  // Set top row to bold
  var topRow = summarySheet.getRange(1, 1, 1, summarySheet.getLastColumn());
  topRow.setFontWeight("bold");
  topRow.setNumberFormat("$00.00");

  // Set bottom row to bold
  var bottomRow = summarySheet.getRange(summarySheet.getLastRow(), 1, 1, summarySheet.getLastColumn());
  bottomRow.setNumberFormat("$00.00");
  bottomRow.setFontWeight("bold");

  scheduleSheet.getRange(scheduleSheet.getDataRange().getLastRow(), 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).setFontWeight("bold");

  let targetRange = scheduleSheet.getRange(bottomRowToUse, 2, 3, summarySheet.getLastColumn());
  var rangeName = scheduleSheetName + magic_namedRangeIdentifier + magic_RangeTypeInitiativeCosts;
  SpreadsheetApp.getActive().setNamedRange(rangeName, targetRange);
}

function populateInitiativeCostPerSprint(summarySheetName, initiativesObj, sprintObj) {
  let summarySheet = SpreadsheetApp.getActive().getSheetByName(summarySheetName);
  var summarySheetLastColumn = summarySheet.getLastColumn() + 2;

  for (var sprintCounter = 0; sprintCounter < Object.keys(sprintObj).length; sprintCounter++) {

    var thisSprint = Object.values(sprintObj)[sprintCounter];
    var sprintName = thisSprint.Name;

    let thisSprintInitiativeCosts = Array();

    for (var initiativeCounter = 0; initiativeCounter < Object.keys(initiativesObj).length; initiativeCounter++) {
      var backgroundColour = Object.keys(initiativesObj)[initiativeCounter];
      var initiative = initiativesObj[backgroundColour];
      var initiativeName = initiative.Name;

      if (sprintCounter == 0) {
        if (initiativeCounter == 0) {
          thisSprintInitiativeCosts.push(magic_HeadingInitiatives);
        }
        thisSprintInitiativeCosts.push(initiativeName);
      } else {
        if (initiativeCounter == 0) {
          thisSprintInitiativeCosts.push(sprintName);
        }

        let costOfThisInitiativeThisSprint = thisSprint.Initiatives.find(c => c.BackgroundColour == backgroundColour);
        if (costOfThisInitiativeThisSprint !== undefined) {
          thisSprintInitiativeCosts.push(costOfThisInitiativeThisSprint.Cost);
        } else {
          thisSprintInitiativeCosts.push(0);

        }
      }
    }

    thisSprintInitiativeCosts.push(thisSprint.Days);
    thisSprintInitiativeCosts.push(thisSprint.Cost);

    summarySheet.getRange(1, summarySheetLastColumn, thisSprintInitiativeCosts.length, 1).setValues(getValuesAs2DArray(thisSprintInitiativeCosts));
    summarySheetLastColumn++;
    console.log(JSON.stringify(thisSprintInitiativeCosts));
  }
}

function calculateAllEstimates() {
  var scheduleSheets = getScheduleSheetNames();
  for (var i = 0; i < scheduleSheets.length; i++) {
    var sheet = scheduleSheets[i];
    calculateDaysPerInitiative(false, false, sheet);
  }
}


function calculateAllCosts() {
  var scheduleSheets = getScheduleSheetNames();
  for (var i = 0; i < scheduleSheets.length; i++) {
    var sheet = scheduleSheets[i];
    calculateDaysPerInitiative(true, false, sheet);
  }
}

function getScheduleSheetNames() {
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get an array of all sheets in the spreadsheet
  var sheets = ss.getSheets();

  // Filter the array to include only sheets that start with "schedule" and do not end with "summary"
  var scheduleSheets = sheets.filter(function (sheet) {
    var name = sheet.getName();
    return name.startsWith(magic_scheduleSheetName) && !name.endsWith(magic_summarySheetName);
  });

  // Get an array of just the sheet names
  var scheduleSheetNames = scheduleSheets.map(function (sheet) {
    return sheet.getName();
  });

  // Log the sheet names
  return scheduleSheetNames;
}


function calculateDaysPerInitiative(doCosts = false, rebuildSheet = false, thisSheetName = null) {

  if (thisSheetName == null) {
    thisSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  }
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
  let summarySheet = createOrDeleteOrUpdateSummarySheet(summarySheetName, initiativesObj, rebuildSheet, scheduleSheetName, sprintObj);
  processSchedule(scheduleSheetName, devsObj, sprintObj, initiativesObj);
  populateSummaryWithEstimates(summarySheetName, initiativesObj, sprintObj, scheduleSheetName);
  summarySheetLastRow = summarySheet.getLastRow();
  updateScheduleWithEstimates(summarySheetName, scheduleSheetName, devsObj, initiativesObj);

  if (doCosts) {
    populateDevDaysPerInitiative(summarySheetName, devsObj, initiativesObj);
    populateInternalAndExternalCosts(summarySheetName, scheduleSheetName, initiativesObj, devsObj);
    populateInitiativeCostPerSprint(summarySheetName, initiativesObj, sprintObj);
    copyInitiativeCostPerSprintToSchedule(summarySheetName, scheduleSheetName, sprintObj);
  }
  formatSummarySheet(summarySheetName, initiativesObj);
}

function copyInitiativeCostPerSprintToSchedule(summarySheetName, scheduleSheetName, sprintObj) {
  let scheduleSheet = SpreadsheetApp.getActive().getSheetByName(scheduleSheetName);
  var range = getInitiativeRange(summarySheetName);
  const values = range.getValues();
  const lastRow = scheduleSheet.getLastRow();
  const targetRow = lastRow + 2;
  var firstColumn = undefined;
  var lastColumn = undefined;

  for (let col = 0; col < values[0].length; col++) {
    const cellValue = values[0][col];
    const sprint = sprintObj.find(s => s.Name == cellValue);
    if (sprint && Object.keys(sprint.Columns).length > 0) {
      const sprintName = sprint.Name;
      const sprintNameRange = scheduleSheet.getRange(targetRow, sprint.Columns[0]); // select the first cell of the range for the sprint name
      sprintNameRange.setValue(sprintName); // set the value of the cell to the sprint name

      const sprintDays = scheduleSheet.getRange(targetRow + 1, sprint.Columns[0]);
      sprintDays.setValue(sprint.Days);

      const sprintCost = scheduleSheet.getRange(targetRow + 2, sprint.Columns[0]);
      sprintCost.setValue(sprint.Cost);
      if (firstColumn === undefined) {
        firstColumn = col;
      }
      lastColumn = col;
    }
  }
  // set the named schedule for the initiative costs by sprint
  var targetRange = scheduleSheet.getRange(targetRow, firstColumn, 3, scheduleSheet.getLastColumn());
  var rangeName = scheduleSheetName + magic_namedRangeIdentifier + magic_RangeTypeSprintCosts;
  SpreadsheetApp.getActive().setNamedRange(rangeName, targetRange);
}

function updateScheduleWithEstimates(summarySheetName, scheduleSheetName, devsObj, initiativesObj) {
  let summarySheet = SpreadsheetApp.getActive().getSheetByName(summarySheetName);
  let scheduleSheet = SpreadsheetApp.getActive().getSheetByName(scheduleSheetName);
  let allDevCount = Object.keys(devsObj).length;
  scheduleSheet.getRange(allDevCount + index_row_scheduleDevNamesBelow + 1, 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).clearContent();
  scheduleSheet.getRange(allDevCount + index_row_scheduleDevNamesBelow + 1, 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).clearFormat();
  let targetRow = allDevCount + index_row_scheduleDevNamesBelow + 2;
  var sourceRange = summarySheet.getRange('A1:D' + Object.keys(initiativesObj).length + 1); // add on 1 because of header row
  transposeDataWithFormat(scheduleSheetName, targetRow, sourceRange, magic_RangeTypeEstimates);
}

function formatSummarySheet(summarySheetName, initiativesObj) {
  let summarySheet = SpreadsheetApp.getActive().getSheetByName(summarySheetName);
  summarySheet.autoResizeColumns(1, summarySheet.getLastColumn());
  var columnIndexes = [index_col_outputCurrentEstimates, index_col_outputTotalInShedule, index_col_outputDifferences, index_col_outputDifferencesFromCurrentSprint];
  applyFormattingToSummarySheet(summarySheetName, initiativesObj, columnIndexes);
}

function logSprintDetails(columnSprint) {
  console.log(`Object: ${JSON.stringify(columnSprint)}`);
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

function transposeDataWithFormat(scheduleSheetName, targetRow, sourceRange, rangeType) {

  var currentSprintCol = getCurrentDateRangeColumn(scheduleSheetName);
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(scheduleSheetName); // Get the target sheet by name

  var sourceData = sourceRange.getValues(); // Get data from range A:F in the source sheet
  var sourceNumberFormat = sourceRange.getNumberFormats(); // Get number formats from range A:F in the source sheet

  var transposedData = transposeArray(sourceData); // Transpose the data  
  var transposedNumberFormat = transposeArray(sourceNumberFormat); // Transpose the number formats

  var rowsOfData = transposedData.length;
  var colsOfData = transposedData[0].length;

  // paste the summaries into the current sprintz
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
  var rangeName = scheduleSheetName + magic_namedRangeIdentifier + rangeType;
  SpreadsheetApp.getActive().setNamedRange(rangeName, targetRange);
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

function getCurrentSprintObjByDate(sprintObj) {
  const currentDate = new Date();
  var currentSprintObject = sprintObj.find(sp => (new Date(sp.Start)).getTime() <= currentDate && (new Date(sp.End)).getTime() >= currentDate);
  return currentSprintObject;
}

function createSummarySheet(summarySheetName, initiativesObj, scheduleSheetName, sprintObj) {

  var currentSprintObject = getCurrentSprintObjByDate(sprintObj);
  summarySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  summarySheet.setName(summarySheetName);
  summarySheet.getRange(1, index_col_outputInitiatives, 1, 1).setValue(magic_HeadingInitiatives);
  summarySheet.getRange(1, index_col_outputCurrentEstimates, 1, 1).setValue(magic_HeadingCurrentEstimates);
  summarySheet.getRange(1, index_col_outputTotalInShedule, 1, 1).setValue(magic_HeadingTotalInSchedule);
  summarySheet.getRange(1, index_col_outputDifferences, 1, 1).setValue(magic_HeadingDifferences);
  summarySheet.getRange(1, index_col_outputEstimateRemaining, 1, 1).setValue(magic_HeadingEstimateRemaining);
  summarySheet.getRange(1, index_col_outputTotalInSheduleFromCurrentSprint, 1, 1).setValue(magic_HeadingTotalInScheduleFromCurrentSprint + " " + currentSprintObject.Name);
  summarySheet.getRange(1, index_col_outputDifferencesFromCurrentSprint, 1, 1).setValue(magic_HeadingTotalInScheduleFromCurrentSprint + " " + currentSprintObject.Name + " " + magic_HeadingDifferences);
  summarySheet.getRange(1, index_col_outputSprints, 1, 1).setValue(magic_HeadingSprints);
  summarySheet.getRange(1, index_col_outputInitiativeCompleteSprint, 1, 1).setValue(magic_HeadingInitiativeCompletedInSprints);
  updateSummarySheet(summarySheetName, initiativesObj, sprintObj);
  var columnIndexes = [index_col_outputDifferences];
  applyFormattingToSummarySheet(summarySheetName, initiativesObj, columnIndexes);
  return summarySheet;
}

function updateSummarySheet(summarySheetName, initiativesObj, sprintObj) {
  var currentSprintObject = getCurrentSprintObjByDate(sprintObj);
  const summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheetName);
  var initiativeObjects = Object.values(initiativesObj);

  const currentEstimates = initiativeObjects.map(initiative => [initiative.CurrentEstimate]);
  const initiativeNames = initiativeObjects.map(initiative => [initiative.Name]);

  summarySheet.getRange(2, index_col_outputInitiatives, Object.keys(initiativesObj).length, 1).setValues(initiativeNames);
  summarySheet.getRange(2, index_col_outputCurrentEstimates, Object.keys(initiativesObj).length, 1).setValues(currentEstimates);
  summarySheet.getRange(1, index_col_outputEstimateRemaining, 1, 1).setValue(magic_HeadingEstimateRemaining + " " + currentSprintObject.Name);
  summarySheet.getRange(1, index_col_outputTotalInSheduleFromCurrentSprint, 1, 1).setValue(magic_HeadingTotalInScheduleFromCurrentSprint + " " + currentSprintObject.Name);
  summarySheet.getRange(1, index_col_outputDifferencesFromCurrentSprint, 1, 1).setValue(magic_HeadingTotalInScheduleFromCurrentSprint + " " + currentSprintObject.Name + " " + magic_HeadingDifferences);
  return summarySheet;
}

function applyFormattingToSummarySheet1(summarySheetName, initiativesObj, columnIndexes) {
  var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheetName);
  var initiativeObjects = Object.values(initiativesObj);
  const bgColours = initiativeObjects.map(initiative => [initiative.BackgroundColour]);
  for (let bg = 0; bg < bgColours.length; bg++) {
    summarySheet.getRange(bg + 2, index_col_outputInitiatives, 1, 1).setBackgrounds(bgColours);
    summarySheet.getRange(bg + 2, index_col_outputDifferences, 1, 1).setNumberFormat('0.00%');
  }

  for (let i = 0; i < columnIndexes.length; i++) {
    var rangeA1Notation = getColumnRange(summarySheetName, columnIndexes[i]);
    applyConditionalFormatting(index_col_outputDifferences, rangeA1Notation, summarySheet);
  }

  return summarySheet;
}

function applyFormattingToSummarySheet(summarySheetName, initiativesObj, columnIndexes) {
  const summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheetName);
  const initiativeObjects = Object.values(initiativesObj);

  const bgColors = initiativeObjects.map(initiative => [initiative.BackgroundColour]);
  const numFormat = [['0.00%']];

  for (let i = 0; i < initiativeObjects.length; i++) {
    const row = i + 2;
    summarySheet.getRange(row, index_col_outputInitiatives, 1, 1).setBackground(bgColors[i]);
    summarySheet.getRange(row, index_col_outputDifferences, 1, 1).setNumberFormat(numFormat);
    summarySheet.getRange(row, index_col_outputDifferencesFromCurrentSprint, 1, 1).setNumberFormat(numFormat);
  }

  for (let i = 0; i < columnIndexes.length; i++) {
    const rangeA1Notation = getColumnRange(summarySheetName, columnIndexes[i]);
    applyConditionalFormatting(index_col_outputDifferences, rangeA1Notation, summarySheet);
  }

  // make the top row bold
  summarySheet.getRange(1, 1, 1, summarySheet.getLastColumn()).setFontWeight("bold");

  return summarySheet;
}
function getParsedDates(sprintObj) {
  const dateFormatRegex = /(\w{3}) (\d{1,2}) (\w{3}) (\d{4})/;
  const parsedArr = sprintObj.map(subArr => {
    return subArr.map(element => {
      if (typeof element === 'string' && element.match(dateFormatRegex)) {
        const standardDateString = element.replace(dateFormatRegex, '$2 $3 $4');
        return new Date(standardDateString);
      }
      return element;
    });
  });
  return parsedArr;
}
function getCurrentDateRangeColumn(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  var currentDate = new Date();  // Get current date
  var startRow = index_row_scheduleSprintStartDateRow;
  var endRow = index_row_scheduleSprintEndDateRow;

  var values = sheet.getRange(startRow, 1, endRow, sheet.getLastColumn()).getValues();
  const nonempty = getParsedDates(values);

  for (var col = 0; col < nonempty[0].length; col++) {
    var startDate = new Date(nonempty[0][col]);
    var endDate = new Date(nonempty[1][col]);

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

function createOrDeleteOrUpdateSummarySheet(summarySheetName, initiativesObj, rebuildSheet, scheduleSheetName, sprintObj) {
  let summarySheet = SpreadsheetApp.getActive().getSheetByName(summarySheetName);

  if (summarySheet != null && rebuildSheet == true) {
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(summarySheet);
    summarySheet = null;
  }

  if (summarySheet == null) {
    summarySheet = createSummarySheet(summarySheetName, initiativesObj, scheduleSheetName, sprintObj);
  }
  else {
    summarySheet.getRange(1, 5, summarySheet.getLastRow(), summarySheet.getLastColumn()).clearContent();
    updateSummarySheet(summarySheetName, initiativesObj, sprintObj);
  }
  return summarySheet;
}
