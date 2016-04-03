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

  var totalRows = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var headerRow = getHeaderRow(sheet, totalRows, lastColumn);
  var bodyRows = getBodyRows(sheet, totalRows, lastColumn);

  setFontFamilies(sheet, headerRow, bodyRows);
  setFontLines(sheet, headerRow, bodyRows);
  setFontSizes(sheet, headerRow, bodyRows);
  setFontWeights(sheet, headerRow, bodyRows);
  setHorizontalAlignments(sheet, headerRow, bodyRows);

  autoResizeColumns(sheet, lastColumn);
}

function setFontFamilies(sheet, headerRow, bodyRows) {
  if (headerRow) headerRow.setFontFamily(headerPreferences.fontFamily);
  if (bodyRows) bodyRows.setFontFamily(bodyPreferences.fontFamily);
};

function setFontLines(sheet, headerRow, bodyRows) {
  if (headerRow) headerRow.setFontLine(headerPreferences.fontLine);
  if (bodyRows) bodyRows.setFontLine(bodyPreferences.fontLine);
}

function setFontSizes(sheet, headerRow, bodyRows) {
  if (headerRow) headerRow.setFontSize(headerPreferences.fontSize);
  if (bodyRows) bodyRows.setFontSize(bodyPreferences.fontSize);
};

function setFontWeights(sheet, headerRow, bodyRows) {
  if (headerRow) headerRow.setFontWeight(headerPreferences.fontWeight);
  if (bodyRows) bodyRows.setFontWeight(bodyPreferences.fontWeight);
};

function setHorizontalAlignments(sheet, headerRow, bodyRows) {
  if (headerRow) headerRow.setHorizontalAlignment(headerPreferences.alignment);
  if (bodyRows) bodyRows.setHorizontalAlignment(bodyPreferences.alignment);
};

function autoResizeColumns(sheet, lastColumn) {
  for (i = 1; i <= lastColumn; i++) {
    sheet.autoResizeColumn(i);
  }
}

function getHeaderRow(sheet, totalRows, lastColumn) {
  if (totalRows > 0) return sheet.getRange(1, 1, 1, lastColumn);  
  return null;
}
function getBodyRows(sheet, totalRows, lastColumn) {
  if (totalRows > 1) return sheet.getRange(2, 1, totalRows - 1, lastColumn);
  return null;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Beautify')
      .addItem('Active sheet', 'main')
      .addToUi();
}
