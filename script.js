// ==================================
// RepairFormat for Google Sheets 1.0
// ==================================
//
// INSTALLATION
// ------------
// 
// If you haven't installed this as an add-on, you can use it by clicking
// 'Tools' > 'Script editor' in Sheets and then pasting the code into the
// Code.gs document. You might then need to "bind" the script to the sheet
// you're working with.
//
// USAGE
// -----
//
// When you click on the add-on's menu item, it will perform the following
// operations on every part of the sheet which has content:
//
// 1. Clears formatting for text (bold, italics, text color, etc.)
// 2. Sets 'Font size'
// 3. Sets 'Font' (font family)
// 4. Sets 'Horizontal align'
// 5. Auto-resizes all columns to fit their content (has a slight delay)

var preferences = {
  fontSize: 14,
  alignment: "left",
  fontFamily: "Consolas"
};

function repairFormat() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clearFormats();
  setFontSize(sheet);
  setFontFamily(sheet);
  setHorizontalAlignment(sheet);
  autoResizeColumns(sheet);
}

function setFontSize(sheet) {
  sheet.getDataRange().setFontSize(preferences.fontSize);
};

function setFontFamily(sheet) {
  sheet.getDataRange().setFontFamily(preferences.fontFamily);
};

function setHorizontalAlignment(sheet) {
  sheet.getDataRange().setHorizontalAlignment(preferences.alignment);
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