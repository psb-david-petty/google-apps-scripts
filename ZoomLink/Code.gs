/**
 * Create 'Zoom Menu' w/ 'Replace Zoom Links' item.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Zoom Menu')
      .addItem('Replace Zoom Links', 'zoom')
      .addToUi();
}

const rules = [
  // start   color    background
    [ 'Z', '#000000', '#ffff00', ],  // yellow (0)
    [ 'A', '#000000', '#cc4125', ],  // light red berry 1 (3)
    [ 'B', '#efefef', '#000099', ],  // custom Navy
    [ 'C', '#000000', '#c27ba0', ],  // light magenta 1 (3)
    [ 'D', '#000000', '#e69138', ],  // dark orange 1 (4)
    [ 'E', '#000000', '#6aa84f', ],  // dark green 1 (4)
    [ 'F', '#000000', '#6d9eeb', ],  // light cornflower blue 1 (3)
    [ 'G', '#000000', '#8e7cc3', ],  // light purple 1 (3)
    [ 'T', '#000000', '#ffffff', ],  // white
    [ 'X', '#000000', '#ffffff', ],  // white
    [ 'L', '#000000', '#cccccc', ],  // gray
    [ 'P', '#000000', '#ffffff', ],  // white
];

/**
 * Return color for text starting with start, or '#000000' if none do.
 * @param {string} text - text to compare with start
 * @return {string} color for text starting with start, or '#000000'if none do
 */
function color(text) {
  for (let [ start, color, background, ] of rules) {
    if (text.startsWith(start)) {
      return color;
    }
  }
  return '#000000';
}

/**
 * Modify the selection as follows.
 * - Update the rich text for each non-empty cell
 * - Replace a Zoom URI with a link to it (match color, no underline)
 *
 * Also, set BHS schedule conditional formatting rules for entire sheet.
 */
function zoom() {
  // https://mashe.hawksey.info/2020/04/everything-a-google-apps-script-developer-wanted-to-know-about-reading-hyperlinks-in-google-sheets-but-was-afraid-to-ask/
  // https://gist.github.com/tanaikech/d39b4b5ccc5a1d50f5b8b75febd807a6
  const replacement = 'Zoom';

  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selection = activeSheet.getSelection();
  Logger.log(`Active Sheet: ${selection.getActiveSheet().getName()} `
    + `Active Range(s): ${selection.getActiveRangeList().getRanges()
      .map(r => r.getA1Notation())}`);

  // Handle discontinuous ranges.
  // TODO: if sheet is filled by IMPORTRANGE, subsequent discontinuous ranges are blank
  var ranges =  selection.getActiveRangeList().getRanges();
  for (let range of ranges) {
    let rangeWidth, rangeHeight, firstColumn, firstRow, lastColumn, lastRow;
    [ rangeWidth, rangeHeight, firstColumn, firstRow, lastColumn, lastRow, ]
      = [ range.getWidth(), range.getHeight(), range.getColumn(), range.getRow(), 
         range.getLastColumn(), range.getLastRow(), ];
    Logger.log(`Range: ${range.getA1Notation()}: `
      + `${firstColumn} - ${lastColumn} = ${rangeWidth} wide; `
      + `${firstRow} - ${lastRow} = ${rangeHeight} high;`)

    // Get values of range to search.
    let rangeValues = activeSheet
      .getRange(firstRow, firstColumn, lastRow, lastColumn)
      .getValues();
    for (let i = 0; i < rangeWidth; i++) {
      for (let j = 0; j < rangeHeight; j++) {
        // Note: Use toString() to avoid type errors on values.
        let value = rangeValues[j][i].toString();
        if (value) Logger.log(`col:${i} row:${j} val:"${value}"`);

        // Only modify non-empty cells.
        if (value) {
          let cell = activeSheet.getRange(firstRow + j, firstColumn + i);
          // Default cell rich text is value.
          let richTextValue = SpreadsheetApp.newRichTextValue()
            .setText(value)
            .build()
          // Match Zoom URI, e.g.:
          // https://zoom.us/j/9726336948?pwd=MkZWNm1XeldQNetwcUpTOGkySWY1Zz09
          // https://psbma-org.zoom.us/j/98961221343?pwd=b1ZZTnB4aHd4YlV1dVNXL2VxMjNuQT09
          let regExp = new RegExp('http[s:/]{3,4}[.-0-9a-zA-Z_]*zoom[.]us[-/?=0-9a-zA-Z_]*', 'gi');
          let match = regExp.exec(value);
          if (match) {
            let uri = match[0];
            let index = value.indexOf(uri);
            Logger.log(`match: "${value}" "${uri}" @ ${index}`);
            
            // Change text style for link: match color, no underline.
            let textStyle = SpreadsheetApp.newTextStyle()
              .setForegroundColor(color(value))
              .setUnderline(false)
              .build();
            // Replace URI string w/ repacement.
            richTextValue = SpreadsheetApp.newRichTextValue()
              .setText(value.replace(uri, replacement))
              .setLinkUrl(index, index + replacement.length, uri)
              .setTextStyle(index, index + replacement.length, textStyle)
              .build();
          }
          // Update cell w/ new rich text.
          cell.setRichTextValue(richTextValue);
        }
      }
    }
  }
  // Set BHS schedule conditional formatting rules. 
  setScheduleConditionalFormattingRules(activeSheet);
}

/**
 * Set BHS schedule conditional formatting rules.
 */
function setScheduleConditionalFormattingRules(sheet) {
  let [ numRows, numCols, ] = [ sheet.getLastRow(), sheet.getLastColumn(), ];
  // TODO: do not apply to row 1 - whenTextStartsWith('F') includes 'Friday'
  let range = sheet.getRange(2, 1, numRows, numCols);
  let conditionalFormatRules = sheet.getConditionalFormatRules();
  for (let [ start, color, background, ] of rules) {
    Logger.log(`Rule: ${start} ${color} ${background}`);
    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith(start)
      .setFontColor(color)
      .setBackground(background)
      .setRanges([range])
      .build();
    conditionalFormatRules.unshift(rule);  // add rule first
  }
  sheet.setConditionalFormatRules(conditionalFormatRules);
}
// 4567890123456789012345678901234567890123456789012345678901234567890
