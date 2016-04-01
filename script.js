// ----------------------------------
// RepairFormat for Google Sheets 1.0
// ----------------------------------
//
// INSTALLATION:
// 
// If you haven't installed this as an add-on, you can use it by clicking
// 'Tools' > 'Script editor' in Sheets and then pasting the code into the
// Code.gs document. You might then need to "bind" the script to the sheet
// you're working with.
//12345678901234567890123456789012345678901234567890123456789012345678901234567|
//12345678901234567890123456789012345678901234567890123456789012345678901234567|
//12345678901234567890123456789012345678901234567890123456789012345678901234567|
//12345678901234567890123456789012345678901234567890123456789012345678901234567|
//12345678901234567890123456789012345678901234567890123456789012345678901234567|
//
// USAGE:
//
// When you click on the add-on's menu item, it will perform the following
// operations on every part of the sheet which has content:
//
// 1. Clear formatting for text (bold, italics, text color, etc.)
// 2. Set 'Font size' to a consistent value across all content
// 3. Auto-resize all columns to fit their content
// 4. Set 'Horizontal align' to 'Left'
//

fontSize = 14;
alignment = "left";

function repairFormat() {
  var sheet = SpreadsheetApp.getActiveSheet();
  clearFormats(sheet);
  increaseFontSize(sheet);
  horizontallyAlign(sheet);
  autoResizeColumns(sheet);
}

function clearFormats(sheet) {
  sheet.clearFormats();
}

function increaseFontSize(sheet) {
  sheet.getDataRange().setFontSize(fontSize);
};

function horizontallyAlign(sheet) {
  sheet.getDataRange().setHorizontalAlignment(alignment);
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