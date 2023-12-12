function colorRowsByMonth() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange("I1:I"); // Column with days
  var values = range.getValues();
  var sum = 0;
  var startRow = 2; // Start from row 2 to skip header
  var month = 1; // Start with month 1

  Logger.log("Starting to process rows");

  for (var i = 1; i < values.length; i++) { // Start from the second row
    var dayValue = parseFloat(values[i][0]);

    if (!isNaN(dayValue) && values[i][0] !== "") {
      sum += dayValue;
      Logger.log("Row: " + (i + 1) + ", Days: " + dayValue + ", Running Total: " + sum);
    }

    // Check if sum exceeds 30 or it's the last row
    if (sum >= 30 || i === values.length - 1) {
      var color = generateRandomColor();
      sheet.getRange(startRow, 10, i - startRow + 1).setBackground(color); // Coloring Column J

      // Set month numbers only for non-empty rows
      for (var j = startRow; j <= i; j++) {
        if (sheet.getRange(j, 9).getValue() !== "") {
          sheet.getRange(j, 10).setValue(month);
        }
      }

      Logger.log("Coloring rows " + startRow + " to " + (i + 1) + " with color " + color + " and month " + month);

      sum = 0;
      startRow = i + 1;
      if (i < values.length - 1) { // Increment month only if it's not the last row
        month++;
      }
    }
  }

  Logger.log("Finished processing rows");
}

function generateRandomColor() {
  // Generate lighter colors by picking random values close to 255
  var red = 200 + Math.floor(Math.random() * 55); // 200 to 255
  var green = 200 + Math.floor(Math.random() * 55); // 200 to 255
  var blue = 200 + Math.floor(Math.random() * 55); // 200 to 255
  return rgbToHex(red, green, blue);
}

function rgbToHex(r, g, b) {
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}
