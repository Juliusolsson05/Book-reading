function colorRowsByMonth() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var daysRange = sheet.getRange("I2:I" + sheet.getLastRow()); // Column with days
  var readingStatusRange = sheet.getRange("M2:P" + sheet.getLastRow()); // Columns M and P for reading status
  var daysValues = daysRange.getValues();
  var readingStatusValues = readingStatusRange.getValues();
  var estimatedTimeRange = sheet.getRange("J2:J" + sheet.getLastRow()); // Column for estimated time
  var sum = 0;
  var month = 1; // Start with month 1
  var startUpcomingMonthRow = -1;

  // Calculate and set estimated reading time
  for (var i = 0; i < daysValues.length; i++) {
    if (daysValues[i][0] !== "") {
      var estimatedTime = parseFloat(daysValues[i][0]) / 3;
      estimatedTimeRange.getCell(i + 1, 1).setValue(estimatedTime); // +1 because range starts from row 2
    }
  }

  Logger.log("Assigning month numbers based on clustering");

  // Assign month numbers based on clustering
  for (var i = 1; i < daysValues.length; i++) {
    var dayValue = parseFloat(daysValues[i][0]);

    if (!isNaN(dayValue) && daysValues[i][0] !== "") {
      sum += dayValue / 3; // Divide by 3 for the reading time per day

      if (sheet.getRange(i + 1, 9).getValue() !== "") {
        sheet.getRange(i + 1, 11).setValue(month); // Set month number in Column K
      }

      // Find the start of the upcoming month's calculation
      if (startUpcomingMonthRow === -1 && readingStatusValues[i][0] === false && readingStatusValues[i][3] === false) {
        startUpcomingMonthRow = i + 1; // +1 because we start from the next row
      }

      if (sum >= 30 || i === daysValues.length - 1) {
        sum = 0;
        if (i < daysValues.length - 1) {
          month++;
        }
      }
    }
  }

  Logger.log("Calculating upcoming month's reads");

  // Calculate upcoming month's reads
  if (startUpcomingMonthRow !== -1) {
    sum = 0;
    month = 1; // Reset month for upcoming reads

    for (var i = startUpcomingMonthRow; i < daysValues.length; i++) {
      var dayValue = parseFloat(daysValues[i][0]);

      if (!isNaN(dayValue) && daysValues[i][0] !== "") {
        sum += dayValue / 3; // Divide by 3 for the reading time per day

        sheet.getRange(i + 1, 12).setValue(month); // Set month number in Column L for upcoming reads

        if (sum >= 30 || i === daysValues.length - 1) {
          sum = 0;
          month++;
        }
      } else {
        // Fill skipped rows with black color
        sheet.getRange(i + 1, 12).setBackground('black');
      }
    }
  }

  Logger.log("Creating color map for each unique month number");

  // Create a color map for each unique month number
  var colorMap = createColorMap(sheet, "K2:K" + sheet.getLastRow());

  Logger.log("Coloring rows based on month numbers");

  // Color rows based on month numbers
  colorRowsByMonthNumber(sheet, colorMap, "K2:K" + sheet.getLastRow(), 11);
  colorRowsByMonthNumber(sheet, colorMap, "L2:L" + sheet.getLastRow(), 12);

  Logger.log("Finished processing rows");
}

function createColorMap(sheet, range) {
  var monthRange = sheet.getRange(range);
  var monthValues = monthRange.getValues();
  var colorMap = {};

  monthValues.forEach(function (value) {
    if (value[0] && !colorMap[value[0]]) {
      colorMap[value[0]] = generateRandomColor(); // Assign a new color if not already present
    }
  });

  return colorMap;
}

function colorRowsByMonthNumber(sheet, colorMap, range, column) {
  var monthRange = sheet.getRange(range);
  var monthValues = monthRange.getValues();

  monthValues.forEach(function (value, index) {
    if (value[0]) {
      sheet.getRange(index + 2, column).setBackground(colorMap[value[0]]);
    }
  });
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
