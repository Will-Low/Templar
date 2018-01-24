/////////////////////////////////////////////
//// Gather Y category list from the logic sheet
/////////////////////////////////////////////
function gatherYTypes(logicSheetRange, logicSheetValues){
  var yListString = "";
  var yListArray = [];
  var numRows = logicSheetRange.getNumRows();
  for (j = 1; j < numRows; j++){
    if (logicSheetValues[j][1] != ""){
      yListString = yListString + logicSheetValues[j][1] + "\n";
      yListArray.push(logicSheetValues[j][1]);
    }
  }
  return [yListString, yListArray];
}


/////////////////////////////////////////////
//// Gather X category list from the logic sheet
/////////////////////////////////////////////
function gatherXTypes(logicSheetRange, logicSheetValues, yCoord){
  var xListString = "";
  var xListArray = [];
  var numCols = logicSheetRange.getNumColumns();
  for (j = 2; j < numCols; j++){
    if (logicSheetValues[0][j] != ""){
      xListArray.push(logicSheetValues[0][j]);
      if ((yCoord != "") && (logicSheetValues[yCoord + 1][j] != "")){
        xListString = xListString + logicSheetValues[0][j] + "\n";
      }
    }
  }
  return [xListString, xListArray];
}


/////////////////////////////////////////////
//// Append text to the base document
/////////////////////////////////////////////
function append(base, textToAppend){

  base = DocumentApp.openById(base).getBody();
  textToAppend = DocumentApp.openByUrl(textToAppend).getBody();

  var totalElements = textToAppend.getNumChildren();
  for (i = 0; i < totalElements; i++){
    var element = textToAppend.getChild(i).copy();
    var type = element.getType();
    if (type == DocumentApp.ElementType.PARAGRAPH){
      base.appendParagraph(element);
    }
    else if (type == DocumentApp.ElementType.TABLE){
      base.appendTable(element);
    }
    else if (type == DocumentApp.ElementType.LIST_ITEM){
      base.appendListItem(element);
    }
    else{
      throw new Error("According to the doc this type couldn't appear in the body: " + type);
    }
  }

  var listLength = base.getListItems().length;
  return base;
}


/////////////////////////////////////////////
//// Return URL of a template key number
/////////////////////////////////////////////
function getUrl(templateNumber){

  var keySheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[2].getDataRange();
  var keySheetValues = keySheet.getValues();
  var keySheetNumOfRows = keySheet.getNumRows();

  for (k = 1; k < keySheetNumOfRows; k++){
    if (templateNumber == keySheetValues[k][0]){
      var url = keySheet.getCell(k + 1, 2).getFormula().substring(12, keySheet.getCell(k + 1, 2).getFormula().indexOf('","'));
      return url;
    }
  }
}


/////////////////////////////////////////////
//// Set nesting levels for any lists in generated template
/////////////////////////////////////////////
function setNestingLevels(base, listLength){
  for (n = 0; n < listLength; n++){
    var listItems = base.getListItems()[n];
    var nestingLevel = listItems.getNestingLevel();
    if (nestingLevel == 0){
      listItems.setGlyphType(DocumentApp.GlyphType.BULLET);
    }
    else if (nestingLevel == 1){
      listItems.setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
    }
    else if (nestingLevel == 2){
      listItems.setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);
    }
  }
  return;
}


/////////////////////////////////////////////
//// Return document name corresponding to a document key number
/////////////////////////////////////////////
function getName(templateNumber){
  var keySheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[2].getDataRange();
  var keySheetValues = keySheet.getValues();
  var keySheetNumOfRows = keySheet.getNumRows();

  for (k = 1; k < keySheetNumOfRows; k++){
    if (templateNumber == keySheetValues[k][0]){
      return String(keySheet.getCell(k + 1, 1).getValue() + ": " + keySheet.getCell(k + 1, 2).getValue());
    }
  }
}
