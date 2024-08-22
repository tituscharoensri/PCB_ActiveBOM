/////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////// !! DO NOT EDIT BELOW !!  ////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Check Inventory Status', 'checkInventoryStatus')
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
  var bomData = bomSheet.getRange(2, 1, bomSheet.getLastRow() - 1, 3).getValues();

  var errorMessages = [];
  var warningMessages = [];
  var currentTime = new Date(); // Get current time
  
  for (var i = 0; i < bomData.length; i++) {
    if (bomData[i][0] === pcbName) {
      var partNumber = bomData[i][1];
      var quantityUsed = bomData[i][2];
      
      if (inventoryMap.hasOwnProperty(partNumber)) {
        var inventoryRow = inventoryMap[partNumber];
        var availableQuantityCell = inventorySheet.getRange(inventoryRow, 2);
        var availableQuantity = availableQuantityCell.getValue();
        if (availableQuantity >= quantityUsed) {
          availableQuantityCell.setValue(availableQuantity - quantityUsed);
        } else {
          // Record an error if there is not enough quantity
          errorMessages.push([
            currentTime.toLocaleString(), 
            'Not enough quantity for part number ' + partNumber + '. Requested: ' + quantityUsed + ', Available: ' + availableQuantity
          ]);
        }
      } else {
        // Record an error if part number is not found in inventory
        errorMessages.push([
          currentTime.toLocaleString(), 
          'Part number ' + partNumber + ' does not exist in Inventory.'
        ]);
      }
    }
  }

  // Check for warnings: quantities under 10
  var inventoryRows = inventorySheet.getRange(2, 1, inventorySheet.getLastRow() - 1, 2).getValues();
  for (var j = 0; j < inventoryRows.length; j++) {
    var partNumber = inventoryRows[j][0];
    var quantity = inventoryRows[j][1];
    if (quantity < 10) {
      warningMessages.push([
        partNumber, 
        'Quantity is below 10: ' + quantity
      ]);
    }
  }

  // Output error messages with timestamps
  if (errorMessages.length > 0 || warningMessages.length > 0) {
    var html = HtmlService.createHtmlOutput('<h1>Notification</h1>' +
      (errorMessages.length > 0 ? '<h2>Errors</h2><ul>' + 
        errorMessages.map(function(msg) { return '<li>' + msg[0] + ': ' + msg[1] + '</li>'; }).join('') +
        '</ul>' : '') +
      (warningMessages.length > 0 ? '<h2>Warnings</h2><ul>' + 
        warningMessages.map(function(msg) { return '<li>' + msg[0] + ': ' + msg[1] + '</li>'; }).join('') +
        '</ul>' : '')
    )
    .setWidth(400)
    .setHeight(300);
    
    SpreadsheetApp.getUi().showSidebar(html);
  }
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
  var bomData = bomSheet.getRange(2, 1, bomSheet.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < bomData.length; i++) {
    if (bomData[i][0] === pcbName) {
      var partNumber = bomData[i][1];
      var quantityUsed = bomData[i][2];
      
      Logger.log("Processing Part Number: " + partNumber + ", Quantity to Undo: " + quantityUsed);

      if (inventoryMap.hasOwnProperty(partNumber)) {
        var inventoryRow = inventoryMap[partNumber];
        var availableQuantityCell = inventorySheet.getRange(inventoryRow, 2);
        var availableQuantity = availableQuantityCell.getValue();
        Logger.log("Current Available Quantity: " + availableQuantity);
        
        availableQuantityCell.setValue(availableQuantity + quantityUsed);
        
        Logger.log("Updated Available Quantity: " + (availableQuantity + quantityUsed));
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
/////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////  UNDO SUBTRACTION FUNCTIONS FOR EACH PCB  /////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////
function undoSubtractComponentsForCANNode() {
  undoSubtractComponentsForPCB("CANNode");
}
function subtractComponentsForSteeringWheelPCB() {
  subtractComponentsForPCB("SteeringWheelPCB");
}

