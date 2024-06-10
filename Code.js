function saveDataHTML() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName('Sheet1');
  var startTime = new Date().getTime(); // Get the current time in milliseconds

  var delay = 10000; // Delay in milliseconds (10 seconds)
  var retryDelay = 30000; // Delay before retrying the fetch in case of rate limit error (30 seconds)
  var maxRetries = 3; // Maximum number of retries for a fetch request

  // Check if the "isRunning" flag is set
  var flagCell = mainSheet.getRange('Z1'); // Assume the flag is located in cell Z1
  if (flagCell.getValue()) {
    return; // If the flag is set, exit the function
  }

  // Set the "isRunning" flag
  flagCell.setValue(true);

  // Set up to handle where to start processing
  var startRowCell = mainSheet.getRange('Y1'); // Store start row in cell Y1
  var startRow = startRowCell.getValue() || 2; // Default to row 2 if no value set

  try {
    var lastRow = mainSheet.getLastRow();
    var dataRange = mainSheet.getRange("A" + startRow + ":A" + lastRow);
    var data = dataRange.getValues();

    var hasPendingRows = false;

    for (var i = 0; i < data.length; i++) {
      var rowIndex = i + startRow;
      var query = data[i][0].trim();
      if (query === "") continue;

      var url = 'https://www.google.com/search?q=' + encodeURIComponent(query);
      var fetchOptions = {
        'muteHttpExceptions': true,
        'timeoutSeconds': 120 // Set timeout to 120 seconds
      };

      var attempt = 0;
      while (attempt < maxRetries) {
        try {
          Utilities.sleep(delay); // Introduce a delay before each request
          var response = UrlFetchApp.fetch(url, fetchOptions);
          var htmlContent = response.getContentText();
          var imageUrls = extractImageUrls(htmlContent, query);

          if (imageUrls.length === 0) {
            // Handle case where no images are found
            mainSheet.getRange(rowIndex, 3).setValue("No images found");
          } else {
            for (var j = 0; j < imageUrls.length; j++) {
              mainSheet.getRange(rowIndex, 3 + j).setValue(imageUrls[j] ? '=IMAGE("' + imageUrls[j] + '")' : "add manually");
            }
          }

          // Log information in column H
          mainSheet.getRange(rowIndex, 8).setValue("Done - " + new Date());
          SpreadsheetApp.flush(); // Force the spreadsheet to update immediately
          break; // Exit the retry loop on success
        } catch (e) {
          if (e.toString().includes("429")) {
            attempt++;
            Utilities.sleep(retryDelay); // Wait before retrying
            mainSheet.getRange(rowIndex, 8).setValue("Attempt " + attempt + ": Rate limit exceeded, retrying...");
          } else {
            // Handle other types of errors
            for (var j = 0; j < 5; j++) {
              mainSheet.getRange(rowIndex, 3 + j).setValue("add manually");
            }
            mainSheet.getRange(rowIndex, 8).setValue("Error occurred: " + e.toString());
            break; // Exit the retry loop on other errors
          }
        }
      }

      if (Date.now() - startTime > 280000) { // Check if close to the 5 minute execution limit
        hasPendingRows = true;
        startRowCell.setValue(rowIndex + 1); // Save the next row index to start from
        break;
      }
    }

    if (!hasPendingRows) {
      startRowCell.clear(); // Clear the start row if all rows have been processed
    }
  } finally {
    // Clear the "isRunning" flag regardless of the outcome
    flagCell.setValue(false);
  }

  // Log end of execution
  console.log('Execution completed at: ' + new Date());

  // Manage triggers
  manageTriggers();
}

function manageTriggers() {
  var existingTriggers = ScriptApp.getProjectTriggers();
  console.log('Existing triggers count: ' + existingTriggers.length); // Log existing trigger count

  for (var i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() == 'saveDataHTML') {
      ScriptApp.deleteTrigger(existingTriggers[i]);
      console.log('Deleted trigger: ' + i); // Log deleted triggers
    }
  }

  // Create a new trigger that will fire in 4 minutes
  var newTrigger = ScriptApp.newTrigger('saveDataHTML')
    .timeBased()
    .after(100000) // Adjust this time based on typical run duration
    .create();
  console.log('New trigger set: ' + newTrigger.getUniqueId()); // Log new trigger creation
}

function extractImageUrls(htmlContent, query) {
  // First stage: Search for images with the product code
  var strictRegex = new RegExp('imgurl=(https:\\/\\/[^"&]*' + query.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&') + '[^"&]*\\.(jpg|jpeg|png|gif))', 'gi');
  var matches = [];
  var match;
  while ((match = strictRegex.exec(htmlContent)) && matches.length < 5) {
    if (match && match[1]) {
      matches.push(decodeURIComponent(match[1]));
    }
  }

  // Second stage: If no images are found, search for any images
  if (matches.length === 0) {
    var looseRegex = new RegExp('imgurl=(https:\\/\\/[^"&]*\\.(jpg|jpeg|png|gif))', 'gi');
    while ((match = looseRegex.exec(htmlContent)) && matches.length < 5) {
      if (match && match[1]) {
        matches.push(decodeURIComponent(match[1]));
      }
    }
  }

  return matches;
}
