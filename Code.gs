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
let magic_InternalDays = "Internal Days";
let magic_ExternalDays = "External Days";
let magic_ExternalCost = "External Cost";
let magic_LocationInternal = "Internal";

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
  let initiatives = [];

  if (initiativeColumn > 0 && estimatesColumn > 0 && bgColumn > 0) {
    for (let ro = 1; ro < lookupsValues.length; ro++) {

      let initiativeName = lookupsValues[ro][initiativeColumn - 1];
      let initiativeCurrentEstimate = lookupsValues[ro][estimatesColumn - 1];
      let initiativeBgColour = lookupsValues[ro][bgColumn - 1];

      if (initiativeName !== "") {
        let initiative = {
          Name: initiativeName,
          CurrentEstimate: initiativeCurrentEstimate,
          BackgroundColour: initiativeBgColour,
          TotalCost: 0,
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

  sprintObj = setColumnsForSprints(scheduleSheetName, sprintObj);
  let [smallestSprintColumn, largestSprintColumn] = getSmallestLargestSprintColumns(sprintObj);

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
    thisDev.RowInScheduleRange = r;
    console.log("ProcessSchedule: thisDev", JSON.stringify(thisDev));

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

function populateSummaryWithEstimates(summarySheetName, initiativesObj) {
  let summarySheet = SpreadsheetApp.getActive().getSheetByName(summarySheetName);
  var summarySheetLastRow = summarySheet.getLastRow();
  let concatSprintNamesPerInitiative = Array();
  let differences = Array();
  let scheduledDays = Array();

  for (let ro = 2; ro <= summarySheetLastRow; ro++) {
    let initiativeCell = summarySheet.getRange(ro, 1);
    let backgroundColour = initiativeCell.getBackground();
    let sumOfDays = initiativesObj[backgroundColour].Days;
    scheduledDays.push(sumOfDays);

    let estimateCell = summarySheet.getRange(ro, index_col_outputCurrentEstimates);
    let estimate = estimateCell.getValue();
    let difference = ((sumOfDays - estimate) / sumOfDays);
    differences.push(difference);


    var sprintThing = Object.values(initiativesObj[backgroundColour].Sprints);
    concatSprintNamesPerInitiative.push(sprintThing.map(initiative => initiative).join(', '));
  }

  summarySheet.getRange(2, index_col_outputSprints, concatSprintNamesPerInitiative.length, 1).setValues(getValuesAs2DArray(concatSprintNamesPerInitiative));
  summarySheet.getRange(2, index_col_outputTotalInShedule, scheduledDays.length, 1).setValues(getValuesAs2DArray(scheduledDays));
  summarySheet.getRange(2, index_col_outputDifferences, differences.length, 1).setValues(getValuesAs2DArray(differences));
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
    devToInitiative.push();
    summarySheet.getRange(1, summarySheetLastColumn + 1 + devCounter, devToInitiative.length, 1).setValues(getValuesAs2DArray(devToInitiative));
    console.log(JSON.stringify(devToInitiative));
  }
}

function populateInternalAndExternalCosts(summarySheetName, scheduleSheetName, initiativesObj) {
  let summarySheet = SpreadsheetApp.getActive().getSheetByName(summarySheetName);
  var summarySheetLastColumn = summarySheet.getLastColumn();

  let internaldevToInitiative = Array();
  let externaldevToInitiative = Array();
  let externalcostToInitiative = Array();

  internaldevToInitiative.push(magic_InternalDays);
  externaldevToInitiative.push(magic_ExternalDays);
  externalcostToInitiative.push(magic_ExternalCost);

  let totalExternalCost = 0;
  let totalExternalDays = 0;
  let totalInternalDays = 0;

  for (var initiativeCounter = 0; initiativeCounter < Object.keys(initiativesObj).length; initiativeCounter++) {
    var backgroundColour = Object.keys(initiativesObj)[initiativeCounter];
    internaldevToInitiative.push(initiativesObj[backgroundColour].InternalDays);
    externaldevToInitiative.push(initiativesObj[backgroundColour].ExternalDays);
    externalcostToInitiative.push(initiativesObj[backgroundColour].ExternalCost);

    totalExternalCost += initiativesObj[backgroundColour].ExternalCost;
    totalExternalDays += initiativesObj[backgroundColour].ExternalDays;
    totalInternalDays += initiativesObj[backgroundColour].InternalDays;
  }

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

  setFirstCellValue(scheduleSheetName, totalExternalCost);
}

function populateInitiativeCostPerSprint(summarySheetName, scheduleSheetName, initiativesObj, sprintObj)
{
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
          thisSprintInitiativeCosts.push(magic_Initiatives);
        }
        thisSprintInitiativeCosts.push(initiativeName);
      } else {
        if (initiativeCounter == 0) {
          thisSprintInitiativeCosts.push(sprintName);
        }
        else {
          let costOfThisInitiativeThisSprint = thisSprint.Initiatives.find(c => c.BackgroundColour == backgroundColour);
          if (costOfThisInitiativeThisSprint !== undefined) {
            thisSprintInitiativeCosts.push(costOfThisInitiativeThisSprint.Cost);
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
  let summarySheet = createOrDeleteOrUpdateSummarySheet(summarySheetName, initiativesObj, rebuildSheet);
  processSchedule(scheduleSheetName, devsObj, sprintObj, initiativesObj);
  populateSummaryWithEstimates(summarySheetName, initiativesObj, sprintObj);
  summarySheetLastRow = summarySheet.getLastRow();
  populateDevDaysPerInitiative(summarySheetName, devsObj, initiativesObj);
  populateInternalAndExternalCosts(summarySheetName, scheduleSheetName, initiativesObj);
  populateInitiativeCostPerSprint(summarySheetName, scheduleSheetName, initiativesObj, sprintObj);

  throw new Error();






  if (doCosts == true) {

    logSprintsDetails(sprintNamesAndDates);
    updateSprintCosts(allDevCount, devNames, lastScheduleColumn, scheduleRange, sprintNamesAndDates);
    logSprintsDetails(sprintNamesAndDates);

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
  summarySheet.getRange(1, index_col_outputInitiatives, 1, 1).setValue(magic_Initiatives);
  summarySheet.getRange(1, index_col_outputCurrentEstimates, 1, 1).setValue(magic_currentEstimates);
  summarySheet.getRange(1, index_col_outputTotalInShedule, 1, 1).setValue(magic_totalInSchedule);
  summarySheet.getRange(1, index_col_outputDifferences, 1, 1).setValue(magic_Differences);
  summarySheet.getRange(1, index_col_outputSprints, 1, 1).setValue(magic_Sprints);
  updateSummarySheet(summarySheetName, initiativesObj);
  var columnIndexes = [index_col_outputDifferences];
  applyFormattingToSummarySheet(summarySheetName, initiativesObj, columnIndexes);
  return summarySheet;
}

function updateSummarySheet(summarySheetName, initiativesObj) {
  const summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheetName);
  var initiativeObjects = Object.values(initiativesObj);

  const currentEstimates = initiativeObjects.map(initiative => [initiative.CurrentEstimate]);
  const initiativeNames = initiativeObjects.map(initiative => [initiative.Name]);

  summarySheet.getRange(2, index_col_outputInitiatives, Object.keys(initiativesObj).length, 1).setValues(initiativeNames);
  summarySheet.getRange(2, index_col_outputCurrentEstimates, Object.keys(initiativesObj).length, 1).setValues(currentEstimates);
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
  }

  for (let i = 0; i < columnIndexes.length; i++) {
    const rangeA1Notation = getColumnRange(summarySheetName, columnIndexes[i]);
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

