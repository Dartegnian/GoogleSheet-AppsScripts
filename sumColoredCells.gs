function sumColoredCells(countRange, colorRef) {
  var activeRange = SpreadsheetApp.getActiveRange();
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var activeFormula = activeRange.getFormula();

  // Extract the range addresses from the formula
  var countRangeAddress = activeFormula.match(/\((.*)\,/).pop().trim();
  var colorRefAddress = activeFormula.match(/\,(.*)\)/).pop().trim();

  // Get the values and backgrounds in the specified range
  var values = activeSheet.getRange(countRangeAddress).getValues();
  var backgrounds = activeSheet.getRange(countRangeAddress).getBackgrounds();
  
  // Get the background color of the reference cell
  var backgroundColor = activeSheet.getRange(colorRefAddress).getBackground();

  var sumColoredCells = 0;

  for (var i = 0; i < values.length; i++) {
    for (var k = 0; k < values[i].length; k++) {
      // Check if the background color matches the reference color
      if (backgrounds[i][k] == backgroundColor) {
        // If yes, add the corresponding value to the sum
        sumColoredCells += parseFormattedNumber(values[i][k]);
      }
    }
  }

  return sumColoredCells;
}

// Function to parse numbers with commas and handle non-numeric values
function parseFormattedNumber(value) {
  // Check if the value is a string before attempting any string manipulation
  if (typeof value === 'string') {
    // Use a regular expression to extract numerical values
    var numberMatch = value.match(/[-]{0,1}[\d]*[.]{0,1}[\d]+/);
    return numberMatch ? parseFloat(numberMatch[0]) : 0;
  } else {
    // If the value is not a string, return it as is
    return value;
  }
}
