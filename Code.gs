function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    {name: "Export JSON for this sheet", functionName: "exportSheet"},
  ];
  ss.addMenu("Export JSON", menuEntries);
}

function exportSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var rowsData = getRowsData_(sheet);
  var json = makeJSON_(rowsData);
  displayText_(json);
}

function makeJSON_(object) {
  var jsonString = JSON.stringify(object, null, 2);
  return jsonString;
}

function displayText_(text) {
  var output = HtmlService.createHtmlOutput("<textarea style='width:100%;' rows='20'>" + text + "</textarea>");
  output.setWidth(400)
  output.setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(output, 'Exported JSON');
}

function round_(value, digits = 4){
  if (isNaN(value)) return value
  var shouldBeRound = (value.toString().split('.').length === 2) && (value.toString().split('.')[1].length > digits)
  if (!shouldBeRound) return value
  return parseFloat(value.toFixed(digits))
}

function getRowsData_(sheet) {
  var headersRange = sheet.getRange(1, 1, sheet.getFrozenRows(), sheet.getMaxColumns());
  var headers = headersRange.getValues()[0];
  var dataRange = sheet.getRange(sheet.getFrozenRows()+1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  // console.log(dataRange.get)
  var objects = getObjects_(dataRange.getValues(), headers);
  console.log(dataRange.getValues())
  return objects;
}

function getObjects_(data, keys) {
  var objects = {};
  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      var posterStyle = data[i][0].toLowerCase()
      var posterSize = data[i][1]
      if (isCellEmpty_(cellData)) {
        // continue;
      }

      if (posterStyle === '') continue

      if (!objects[posterStyle]){
        objects[posterStyle] = {}
      }

      if (!objects[posterStyle][posterSize]) {
        objects[posterStyle][posterSize] = {}
      }

      if (j === 0) {
        objects[posterStyle][posterSize][keys[j]] = posterStyle
      } else {
        objects[posterStyle][posterSize][keys[j]] = round_(cellData)
      }
    }
  }
  return objects;
}

function isCellEmpty_(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}