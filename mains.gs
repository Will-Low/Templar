// Global vars
var ui = SpreadsheetApp.getUi();
var newDoc = "";
var logicSheetRange = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1].getDataRange();
var logicSheetValues = logicSheetRange.getValues();

function onOpen(){
    ui.createMenu('Templar')
      .addItem('Run with X & Y prompts', 'main')
      .addItem('Generate by Selection', 'cursorMain')
      .addItem('Key Lookup', 'componentPrinter')
      .addItem('Generate by Y Coord', 'byYGeneration')
      .addItem('Generate by X Coord', 'byXGeneration')
      .addToUi();
}


/////////////////////////////////////////////
//// Run with prompts for both X and Y axes.
/////////////////////////////////////////////
function main(){
  var yTypes = gatherYTypes(logicSheetRange, logicSheetValues);
  var ySelection = ui.prompt("Which Y value are you selecting? Options are:\n" + yTypes[0]).getResponseText();
  var yCoord = yTypes[1].indexOf(ySelection);

  var xTypes = gatherXTypes(logicSheetRange, logicSheetValues, yCoord);
  var xSelection = ui.prompt("Which X value are you selecting? Options are:\n" + xTypes[0]).getResponseText();
  var xCoord = xTypes[1].indexOf(xSelection);

  var selectionValue = logicSheetValues[yCoord + 1][xCoord + 2].split(", ");

  newDoc = DriveApp.getFileById(DocumentApp.openByUrl(getUrl(selectionValue[0])).getId()).makeCopy(ySelection + " " + xSelection, DriveApp.getRootFolder()).getId();

  for (m = 1; m < selectionValue.length; m++){
    var base = append(newDoc, getUrl(selectionValue[m]));
  }
  var numListItemsInDocument = base.getListItems().length;
  setNestingLevels(base, numListItemsInDocument);
}


/////////////////////////////////////////////
//// Run based on selected cell.
/////////////////////////////////////////////
function cursorMain(){
  var yTypes = gatherYTypes(logicSheetRange, logicSheetValues);
  var xTypes = gatherXTypes(logicSheetRange, logicSheetValues, 0);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var selectionValue = ss.getActiveSheet().getActiveCell().getValue().split(", ");
  var lastRow = ss.getActiveSheet().getActiveCell().getLastRow();
  var lastColumn = ss.getActiveSheet().getActiveCell().getLastColumn();
  var selectionLocation = [lastRow, lastColumn];
  newDoc = DriveApp.getFileById(DocumentApp.openByUrl(getUrl(selectionValue[0])).getId()).makeCopy(yTypes[1][selectionLocation[0] - 2] + " " + xTypes[1][selectionLocation[1] - 3], DriveApp.getRootFolder()).getId();

  var numOfComponents = selectionValue.length;
  var base;
  for (m = 1; m < numOfComponents; m++){
    base = append(newDoc, getUrl(selectionValue[m]));
  }
  var numListItemsInDocument = base.getListItems().length;
  setNestingLevels(base, numListItemsInDocument);
}


/////////////////////////////////////////////
//// Hold Y constant. Generate for each X.
/////////////////////////////////////////////
function byYGeneration() {
  var continueResponse = ui.alert('This may take a while. You\'ll get a message once complete. \n\nProceed?', ui.ButtonSet.YES_NO);
  if (continueResponse != ui.Button.YES) {
    return;
  }

  var yTypes = gatherYTypes(logicSheetRange, logicSheetValues);

  var ySelection = ui.prompt("Which Y value are you selecting? Options are:\n" + yTypes[0]).getResponseText();
  var yCoord = yTypes[1].indexOf(ySelection);
  if (yCoord == -1){
    ui.alert("Could not find " + ySelection + ".");
    return;
  }

  var xTypes = gatherXTypes(logicSheetRange, logicSheetValues, yCoord);

  var letterListLength = xTypes[1].length;
  for (r = 0; r < letterListLength; r++){
    try {
      var selectionValue = logicSheetValues[yCoord + 1][r + 2].split(", ");
      newDoc = DriveApp.getFileById(DocumentApp.openByUrl(getUrl(selectionValue[0])).getId()).makeCopy(ySelection + " " + xTypes[1][r], DriveApp.getRootFolder()).getId();
      var listLength = selectionValue.length;
      var base;
      for (m = 1; m < listLength; m++){
        base = append(newDoc, getUrl(selectionValue[m]));
      }
      var numListItemsInDocument = base.getListItems().length;
      setNestingLevels(base, numListItemsInDocument);
    }
    catch(err) {
      continue;
    }
  }
  ui.alert("Generation completed");
}


/////////////////////////////////////////////
//// Hold X constant. Generate for each Y.
/////////////////////////////////////////////
function byXGeneration() {
  var continueResponse = ui.alert('This may take a while. You\'ll get a message once complete. \n\nProceed?', ui.ButtonSet.YES_NO);
  if (continueResponse != ui.Button.YES) {
    return;
  }
  var yTypes = gatherYTypes(logicSheetRange, logicSheetValues);
  var yCoord = -1; // Compensates for 0-indexed design of gatherXTypes(). In this case, we are 1-indexed.
  var xTypes = gatherXTypes(logicSheetRange, logicSheetValues, yCoord);
  var xSelection = ui.prompt("Which X value are you selecting? Options are:\n" + xTypes[0]).getResponseText();
  var xCoord = xTypes[1].indexOf(xSelection);
  var numYTypes = yTypes[1].length;

  for (r = 0; r < numYTypes; r++){
      var selectionValue = logicSheetValues[r + 1][xCoord + 2].split(", ");
      Logger.log(selectionValue);
      newDoc = DriveApp.getFileById(DocumentApp.openByUrl(getUrl(selectionValue[0])).getId()).makeCopy(yTypes[1][r] + " " + xSelection, DriveApp.getRootFolder()).getId();
      var base;
      for (m = 1; m < selectionValue.length; m++){
        base = append(newDoc, getUrl(selectionValue[m]));
      }
      var numListItemsInDocument = base.getListItems().length;
      setNestingLevels(base, numListItemsInDocument);
  }
  ui.alert("Generation completed");
}


/////////////////////////////////////////////
//// Print the names of documents that make up a Logic-cell formula
/////////////////////////////////////////////
function componentPrinter(){
  var yTypes = gatherYTypes(logicSheetRange, logicSheetValues);
  var xTypes = gatherXTypes(logicSheetRange, logicSheetValues, 0);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var values = ss.getActiveSheet().getActiveCell().getValue().split(", ");
  var componentList = "";
  for (w = 0; w < values.length; w++){
    componentList = componentList + getName(values[w]) + "\n";
  }
  ui.alert(componentList);
}
