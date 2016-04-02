var headerPreferences = {
  alignment: "center",
  fontFamily: "Consolas",
  fontLine: "underline",
  fontSize: 14,
  fontWeight: "bold"
};

var bodyPreferences = {
  alignment: "left",
  fontFamily: "Consolas",
  fontLine: "none",
  fontSize: 14,
  fontWeight: "normal"
};

function main() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clearFormats();
  setFontFamilies(sheet);
  setFontLines(sheet);
  setFontSizes(sheet);
  setFontWeights(sheet);
  setHorizontalAlignments(sheet);
  autoResizeColumns(sheet);
}

function setFontFamilies(sheet) {
  getHeaderRow(sheet).setFontFamily(headerPreferences.fontFamily);
  getBodyRows(sheet).setFontFamily(bodyPreferences.fontFamily);
};

function setFontLines(sheet) {
  getHeaderRow(sheet).setFontLine(headerPreferences.fontLine);
  getBodyRows(sheet).setFontLine(bodyPreferences.fontLine);
}

function setFontSizes(sheet) {
  getHeaderRow(sheet).setFontSize(headerPreferences.fontSize);
  getBodyRows(sheet).setFontSize(bodyPreferences.fontSize);
};

function setFontWeights(sheet) {
  getHeaderRow(sheet).setFontWeight(headerPreferences.fontWeight);
  getBodyRows(sheet).setFontWeight(bodyPreferences.fontWeight);
};

function setHorizontalAlignments(sheet) {
  getHeaderRow(sheet).setHorizontalAlignment(headerPreferences.alignment);
  getBodyRows(sheet).setHorizontalAlignment(bodyPreferences.alignment);
};

function autoResizeColumns(sheet) {
  lastColumn = sheet.getLastColumn()
  for (i = 1; i < lastColumn; i++) {
    sheet.autoResizeColumn(i);
  }
}

function getBodyRows(sheet) {
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());  
}

function getHeaderRow(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn());  
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Beautify')
      .addItem('Active sheet', 'main')
      .addToUi();
}
