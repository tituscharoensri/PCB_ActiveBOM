/////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////// !! DO NOT EDIT BELOW !!  ////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Check Inventory Status', 'checkInventoryStatus')
    .addItem('Create Usage Graph', 'createUsageGraph') // Added menu item for creating the graph
    .addItem('Highlight Search Results', 'highlightSearchResults')
    .addItem('Reset highlights', 'unhighlightCells')
    .addToUi();
}

function subtractComponentsForPCB(pcbName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var inventorySheet = spreadsheet.getSheetByName('Inventory');
  var bomSheet = spreadsheet.getSheetByName('PCB BOM');

  // Get all inventory data
  var inventoryData = inventorySheet.getRange(2, 1, inventorySheet.getLastRow() - 1, 2).getValues();
  var inventoryMap = {};
  for (var i = 0; i < inventoryData.length; i++) {
    inventoryMap[inventoryData[i][0]] = i + 2;  // Map part number to row number in Inventory
  }

  // Get all BOM data for the specific PCB
  var bomData = bomSheet.getRange(2, 1, bomSheet.getLastRow() - 1, 4).getValues(); // Include 4 columns
  var bomBackgrounds = bomSheet.getRange(2, 1, bomSheet.getLastRow() - 1, 4).getBackgrounds(); // For highlighting

  var errorMessages = [];
  var warningMessages = [];
  var currentTime = new Date(); // Get current time
  var missingPartNumber = false;

  // First pass: Check if all part numbers exist in the Inventory sheet
  for (var i = 0; i < bomData.length; i++) {
    if (bomData[i][0] === pcbName) {
      var partNumber = bomData[i][1];
      if (!inventoryMap.hasOwnProperty(partNumber)) {
        // Record an error if part number is not found in inventory
        errorMessages.push([
          currentTime.toLocaleString(), 
          'Part number ' + partNumber + ' does not exist in Inventory.'
        ]);
        missingPartNumber = true;
      }
    }
  }

  // If a part number is missing, do not proceed with subtraction
  if (missingPartNumber) {
    // Log errors and exit function
    showErrorPopup(errorMessages, warningMessages);
    return;
  }

  // Second pass: Subtract quantities if no part numbers are missing
  for (var i = 0; i < bomData.length; i++) {
    if (bomData[i][0] === pcbName) {
      var partNumber = bomData[i][1];
      var quantityUsed = bomData[i][2];
      
      var inventoryRow = inventoryMap[partNumber];
      var inventoryAvailableQuantityCell = inventorySheet.getRange(inventoryRow, 2);
      var inventoryAvailableQuantity = inventoryAvailableQuantityCell.getValue();
      
      var bomAvailableQuantityCell = bomSheet.getRange(i + 2, 4); // Adjust row index for BOM sheet
      var bomAvailableQuantity = bomAvailableQuantityCell.getValue();
      
      if (inventoryAvailableQuantity >= quantityUsed) {
        // Update Inventory and PCB BOM
        inventoryAvailableQuantityCell.setValue(inventoryAvailableQuantity - quantityUsed);
        bomAvailableQuantityCell.setValue(bomAvailableQuantity - quantityUsed);
        
        // Add a warning if quantity falls below 10
        if (inventoryAvailableQuantity - quantityUsed < 10) {
          warningMessages.push([
            'Warning: Quantity for part number ' + partNumber + ' is low. Remaining: ' + (inventoryAvailableQuantity - quantityUsed)
          ]);
        }
      } else {
        // Record an error if there is not enough quantity
        errorMessages.push([
          currentTime.toLocaleString(), 
          'Not enough quantity for part number ' + partNumber + '. Requested: ' + quantityUsed + ', Available: ' + inventoryAvailableQuantity + '- Only Available Parts will be subtracted!'
        ]);

        // Highlight the row in PCB BOM if available quantity is less than quantity used
        bomBackgrounds[i][0] = '#FF9999'; // Highlight the Available Quantity cell
        bomBackgrounds[i][1] = '#FF9999'; // Highlight the Available Quantity cell
        bomBackgrounds[i][2] = '#FF9999'; // Highlight the Available Quantity cell
        bomBackgrounds[i][3] = '#FF9999'; // Highlight the Available Quantity cell
      }
    }
  }

  // Apply the background color to the BOM sheet
  bomSheet.getRange(2, 1, bomSheet.getLastRow() - 1, 4).setBackgrounds(bomBackgrounds);
  
  // Log errors and warnings to sheet and show popup
  showErrorPopup(errorMessages, warningMessages);
}





function showErrorPopup(errorMessages, warningMessages) {
  var ui = SpreadsheetApp.getUi();
  var message = 'Error Log:\n';
  for (var i = 0; i < errorMessages.length; i++) {
    message += errorMessages[i][0] + ': ' + errorMessages[i][1] + '\n';
  }
  if (warningMessages.length > 0) {
    message += '\nWarnings:\n';
    for (var i = 0; i < warningMessages.length; i++) {
      message += warningMessages[i] + '\n';
    }
  }
  ui.alert('Error and Warning Notification', message, ui.ButtonSet.OK);
}




function undoSubtractComponentsForPCB(pcbName) {
  var inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var bomSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PCB BOM');

  // Logging the start of the process
  Logger.log("Starting undo for PCB: " + pcbName);

  // Get all inventory data
  var inventoryData = inventorySheet.getRange(2, 1, inventorySheet.getLastRow() - 1, 2).getValues();
  var inventoryMap = {};
  for (var i = 0; i < inventoryData.length; i++) {
    inventoryMap[inventoryData[i][0]] = i + 2;  // Map part number to row number in Inventory
    Logger.log("Mapped Part Number: " + inventoryData[i][0] + " to Inventory Row: " + (i + 2));
  }

  // Get all BOM data for the specific PCB
  var bomData = bomSheet.getRange(2, 1, bomSheet.getLastRow() - 1, 4).getValues(); // Include 4 columns
  for (var i = 0; i < bomData.length; i++) {
    if (bomData[i][0] === pcbName) {
      var partNumber = bomData[i][1];
      var quantityUsed = bomData[i][2];
      
      Logger.log("Processing Part Number: " + partNumber + ", Quantity to Undo: " + quantityUsed);

      if (inventoryMap.hasOwnProperty(partNumber)) {
        var inventoryRow = inventoryMap[partNumber];
        var inventoryAvailableQuantityCell = inventorySheet.getRange(inventoryRow, 2);
        var inventoryAvailableQuantity = inventoryAvailableQuantityCell.getValue();
        var bomAvailableQuantityCell = bomSheet.getRange(i + 2, 4); // Adjust row index for BOM sheet
        var bomAvailableQuantity = bomAvailableQuantityCell.getValue();
        
        Logger.log("Current Available Quantity in Inventory: " + inventoryAvailableQuantity);
        
        // Update Inventory and PCB BOM
        inventoryAvailableQuantityCell.setValue(inventoryAvailableQuantity + quantityUsed);
        bomAvailableQuantityCell.setValue(bomAvailableQuantity + quantityUsed);
        
        Logger.log("Updated Available Quantity in Inventory: " + (inventoryAvailableQuantity + quantityUsed));
        Logger.log("Updated Available Quantity in BOM: " + (bomAvailableQuantity + quantityUsed));
      } else {
        Logger.log("Part Number " + partNumber + " not found in Inventory.");
      }
    }
  }
}


function applyConditionalFormatting() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  
  // Define the range for conditional formatting (columns A and B)
  var range = sheet.getRange('A2:B' + sheet.getLastRow());
  
  // Get existing conditional formatting rules
  var rules = sheet.getConditionalFormatRules();
  
  // Clear any existing conditional formatting rules that affect the range
  rules = rules.filter(function(rule) {
    var ruleRanges = rule.getRanges();
    return !ruleRanges.some(function(r) {
      return r.getA1Notation() === range.getA1Notation();
    });
  });

  // Create a new conditional formatting rule to highlight rows where quantity < 10 in column B
  var highlightRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(10)
    .setBackground('#FF9999')  // Light red background for highlighting
    .setRanges([range])
    .build();

  // Apply the updated rules to the sheet
  sheet.setConditionalFormatRules([highlightRule]);
}

function checkInventoryStatus() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var inventorySheet = spreadsheet.getSheetByName('Inventory');
  
  // Get all inventory data
  var inventoryData = inventorySheet.getRange(2, 1, inventorySheet.getLastRow() - 1, 2).getValues();

  var warningMessages = [];
  
  for (var i = 0; i < inventoryData.length; i++) {
    var partNumber = inventoryData[i][0];
    var quantity = inventoryData[i][1];
    if (quantity < 10) {
      warningMessages.push([
        partNumber, 
        'Quantity is below 10: ' + quantity
      ]);
    }
  }

  // Display warning messages in a sidebar
  var html = HtmlService.createHtmlOutput('<h1>Inventory Status</h1>' +
    (warningMessages.length > 0 ? '<h2>Components with Low Quantity</h2><ul>' + 
      warningMessages.map(function(msg) { return '<li>' + msg[0] + ': ' + msg[1] + '</li>'; }).join('') +
      '</ul>' : '<p>All quantities are above or equal to 10.</p>')
  )
  .setWidth(400)
  .setHeight(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

function createUsageGraph() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var bomSheet = spreadsheet.getSheetByName('PCB BOM');
  var inventorySheet = spreadsheet.getSheetByName('Inventory');
  var chartSheet = spreadsheet.getSheetByName('Usage Graph(DO NOT EDIT)');
  
  // Create a new sheet for the chart if it doesn't exist
  if (!chartSheet) {
    chartSheet = spreadsheet.insertSheet('Usage Graph(DO NOT EDIT)');
  } else {
    // Clear existing content if the sheet already exists
    chartSheet.clear();
  }

  // Get BOM data
  var bomData = bomSheet.getRange(2, 1, bomSheet.getLastRow() - 1, 3).getValues(); // Assuming PCB Name in Column A, Part Number in Column B, Quantity Used in Column C
  
  var usageMap = {};

  // Aggregate quantities used for each part number
  for (var i = 0; i < bomData.length; i++) {
    var partNumber = bomData[i][1];
    var quantityUsed = bomData[i][2];
    if (usageMap[partNumber]) {
      usageMap[partNumber] += quantityUsed;
    } else {
      usageMap[partNumber] = quantityUsed;
    }
  }

  // Get inventory data
  var inventoryData = inventorySheet.getRange(2, 1, inventorySheet.getLastRow() - 1, 2).getValues(); // Assuming Part Number in Column A, Available Quantity in Column B
  
  var inventoryMap = {};
  for (var j = 0; j < inventoryData.length; j++) {
    inventoryMap[inventoryData[j][0]] = inventoryData[j][1]; // Map part number to available quantity
  }

  // Prepare data for chart
  var chartData = [['Part Number', 'Total Quantity Used', 'Available Quantity']]; // Header row

  for (var partNumber in usageMap) {
    if (usageMap.hasOwnProperty(partNumber)) {
      var totalQuantityUsed = usageMap[partNumber];
      var availableQuantity = inventoryMap[partNumber] || 0; // Default to 0 if part number not in inventory
      chartData.push([partNumber, totalQuantityUsed, availableQuantity]);
    }
  }
  
  // Add data to the chart sheet
  chartSheet.getRange(1, 1, chartData.length, 3).setValues(chartData);

  // Create the chart
  var dataRange = chartSheet.getRange(1, 1, chartData.length, 3); // Data range includes headers
  var chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    .setPosition(1, 4, 0, 0) // Position the chart
    .setOption('title', 'Part Number Usage (Quantity Required from PCB BOM) vs Available Quantity')
    .setOption('hAxis', {title: 'Part Number'})
    .setOption('vAxis', {title: 'Quantity'})
    .setOption('series', {
      0: {color: 'blue', label: 'Total Quantity Used'},
      1: {color: 'red', label: 'Available Quantity'}
    })
    .setOption('legend', {position: 'top'})
    .build();

  // Insert the chart into the sheet
  chartSheet.insertChart(chart);
  // Add explanation text in a cell near the chart
  chartSheet.getRange('A25').setValue('This chart compares the total quantity of each part number used across PCBs with the available quantity in the inventory').setFontSize(12).setFontColor('black');
    chartSheet.getRange('A26').setValue('Usefull to tell if there are enough components in lab to make a spare of each PCB & which parts are used more frequently accross PCBs').setFontSize(12).setFontColor('black');

}

function highlightSearchResults() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter the text to search for:');
  
  // Check if the user pressed 'Cancel'
  if (response.getSelectedButton() == ui.Button.CANCEL) {
    ui.alert('Search canceled.');
    return;
  }
  
  var searchText = response.getResponseText().trim(); // Trim to remove leading/trailing spaces

  // If searchText is empty, cancel the search
  if (!searchText) {
    ui.alert('No search text entered. Search canceled.');
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();

  // Define the highlight color
  var highlightColor = '#FFFF00'; // Yellow color

  // Get default colors
  var defaultColors = getDefaultColors(sheet);

  // Clear previous highlights (restore the default colors first)
  applyDefaultColors(sheet, defaultColors);

  // Search and highlight cells
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j].toString().indexOf(searchText) !== -1) {
        backgrounds[i][j] = highlightColor;
      }
    }
  }
  
  // Apply the background color to cells with search text
  range.setBackgrounds(backgrounds);
  
  ui.alert('Search completed and highlighted cells.');
}


function unhighlightCells() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();

  // Clear all background colors
  range.setBackground(null);
  
  // Restore default colors
  var defaultColors = getDefaultColors(sheet);
  applyDefaultColors(sheet, defaultColors);


  SpreadsheetApp.getUi().alert('Highlights removed.');
}

function highlightLowStockRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var bomSheet = spreadsheet.getSheetByName('PCB BOM');
  
  // Define the color for highlighting
  var highlightColor = '#FF9999'; // Light red color

  // Get the data from the PCB BOM sheet
  var lastRow = bomSheet.getLastRow();
  var range = bomSheet.getRange(2, 1, lastRow - 1, 4); // Adjust range to include 4 columns (A to D)
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();

  // Loop through the rows to check quantities
  for (var i = 0; i < values.length; i++) {
    var quantityUsed = values[i][2]; // Column C (3rd column)
    var availableQuantity = values[i][3]; // Column D (4th column)
    
    if (availableQuantity < quantityUsed) {
      // Highlight the row if Available Quantity is less than Quantity Used
      for (var j = 0; j < backgrounds[i].length; j++) {
        backgrounds[i][j] = highlightColor;
      }
    }
  }
  
  // Apply the background color to the range
  range.setBackgrounds(backgrounds);
  
  SpreadsheetApp.getUi().alert('Rows with low stock highlighted.');
}


function getDefaultColors(sheet) {
  var colors = {};
  
  if (sheet.getName() === 'Inventory') {
    colors = {
      'A': '#93c47d', // Light Green 1
      'B': '#93c47d', // Light Green 1
      'C': '#93c47d', // Light Green 1
      'D': '#93c47d', // Light Green 1
      'E': '#93c47d', // Light Green 1
      'F': '#ffd966'  // Light Yellow 1
    };
  } else if (sheet.getName() === 'PCB BOM') {
    colors = {
      'A': '#93c47d', // Light Green 1
      'B': '#93c47d', // Light Green 1
      'C': '#93c47d', // Light Green 1
      'D': '#ffd966', // Light Yellow 1
      'E': '#ffd966'  // Light Yellow 1
    };
  }
  
  return colors;
}

function applyDefaultColors(sheet, colors) {
  for (var col in colors) {
    var range = sheet.getRange(col + '1:' + col);
    range.setBackground(colors[col]);
  }
}



function onEdit(e) {
  // Get the active sheet and check if it's the "Inventory" sheet
  var sheet = e.source.getActiveSheet();
  
  // Ensure the script only runs on the "Inventory" sheet
  if (sheet.getName() === "Inventory") {
    // Get the range of the edited cell
    var range = e.range;
    var column = range.getColumn();
    
    // Check if the edited cell is in Column B (quantity column)
    if (column === 2) {
      var row = range.getRow();
      
      // Update the corresponding cell in Column F with the current timestamp
      var timestampCell = sheet.getRange(row, 6);
      timestampCell.setValue(new Date());

      // Set the date-time format to include time to the minute
      timestampCell.setNumberFormat("yyyy-MM-dd HH:mm");
    }
  }
  // update Available Quantities
  var sheet = e.source.getActiveSheet();
  var editedCell = e.range;
  
  // Check if the edit was made in the "Inventory" sheet, specifically in the "Available Quantity" column (B)
  if (sheet.getName() === "Inventory" && editedCell.getColumn() === 2) {
    var partNumber = sheet.getRange(editedCell.getRow(), 1).getValue();
    var newQuantity = editedCell.getValue();
    
    // Get the PCB BOM sheet
    var bomSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PCB BOM');
    var bomData = bomSheet.getRange(2, 2, bomSheet.getLastRow() - 1, 3).getValues(); // Get Part Number and Available Quantity data
    
    // Loop through PCB BOM sheet to find matching part numbers and update their Available Quantity
    for (var i = 0; i < bomData.length; i++) {
      if (bomData[i][0] === partNumber) {
        var bomRow = i + 2; // Adjust for the header row
        bomSheet.getRange(bomRow, 4).setValue(newQuantity);
      }
    }
  }

  // AutoFill PCB BOM Quantity available & Description
  var sheet = e.source.getActiveSheet();
  var editedCell = e.range;

  // Check if the edit was made in the "PCB BOM" sheet, specifically in column B (Part Number)
  if (sheet.getName() === "PCB BOM" && editedCell.getColumn() === 2) {
    var row = editedCell.getRow();
    var partNumber = editedCell.getValue();
    
    // Get the Inventory sheet
    var inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
    var inventoryData = inventorySheet.getRange(2, 1, inventorySheet.getLastRow() - 1, 4).getValues(); // Part Number and Description data
    var inventoryMap = {};
    
    // Create a map for part number to available quantity and description
    for (var i = 0; i < inventoryData.length; i++) {
      var inventoryPartNumber = inventoryData[i][0];
      var availableQuantity = inventoryData[i][1];
      var description = inventoryData[i][3];
      inventoryMap[inventoryPartNumber] = {
        availableQuantity: availableQuantity,
        description: description
      };
    }
    
    // Get the PCB BOM sheet
    var bomSheet = e.source.getSheetByName('PCB BOM');
    
    if (inventoryMap.hasOwnProperty(partNumber)) {
      // Update Available Quantity and Description in PCB BOM for the specific row
      var data = inventoryMap[partNumber];
      bomSheet.getRange(row, 4).setValue(data.availableQuantity); // Column D
      bomSheet.getRange(row, 5).setValue(data.description); // Column E
    } else {
      // Display error notification if part number is not found in Inventory
      showErrorPopup2(partNumber);
      // Clear the entered part number and related columns if part number is not found
      editedCell.setValue('');
      bomSheet.getRange(row, 4).setValue(''); // Clear Column D
      bomSheet.getRange(row, 5).setValue(''); // Clear Column E
    }
  }
}

function showErrorPopup2(partNumber) {
  var ui = SpreadsheetApp.getUi();
  var message = 'Error: Part number ' + partNumber + ' does not exist in the Inventory sheet.\nPlease add it to the Inventory sheet.';
  ui.alert('Part Number Error', message, ui.ButtonSet.OK);
}





/////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////// !! DO NOT EDIT ABOVE !!  ////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////// CAN EDIT BELOW  ////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////  SUBTRACTION FUNCTIONS FOR EACH PCB  //////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////
function subtractComponentsForCANNode() {
  subtractComponentsForPCB("CANNode");
}
function subtractComponentsForSteeringWheelPCB() {
  subtractComponentsForPCB("SteeringWheelPCB");
}
function subtractComponentsForAMI() {
  subtractComponentsForPCB("AMI");
}
/////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////  UNDO SUBTRACTION FUNCTIONS FOR EACH PCB  /////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////
function undoSubtractComponentsForCANNode() {
  undoSubtractComponentsForPCB("CANNode");
}
function undosubtractComponentsForSteeringWheelPCB() {
  undosubtractComponentsForPCB("SteeringWheelPCB");
}

