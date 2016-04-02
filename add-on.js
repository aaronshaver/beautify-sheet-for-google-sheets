var headerRowPreferences = {
  fontSize: 14,
  alignment: "center",
  fontFamily: "Consolas",
  skipHeaderRow: true
};

var sheetPreferences = {
  fontSize: 14,
  alignment: "left",
  fontFamily: "Consolas",
  skipHeaderRow: true
};

function repairFormat() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  clearFormats(sheet);
  setFontSize(sheet);
  setFontFamily(sheet);
  setHorizontalAlignment(sheet);
  autoResizeColumns(sheet);
//  formatHeaderRow();
}

function clearFormats(sheet) {
    sheet.clearFormats();
}

function setFontSize(sheet) {
  sheet.getDataRange().setFontSize(sheetPreferences.fontSize);
};

function setFontFamily(sheet) {
  sheet.getDataRange().setFontFamily(sheetPreferences.fontFamily);
};

function setHorizontalAlignment(sheet) {
  sheet.getDataRange().setHorizontalAlignment(sheetPreferences.alignment);
};

function autoResizeColumns(sheet) {
  lastColumn = sheet.getLastColumn()
  for (i = 1; i < lastColumn; i++) {
    sheet.autoResizeColumn(i);
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('RepairFormat')
      .addItem('Run', 'repairFormat')
      .addToUi();
}