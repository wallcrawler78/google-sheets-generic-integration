/**
 * Rack Configuration Management
 * Handles creation, metadata, and utilities for rack configuration tabs
 */

// Constants for metadata row structure
var METADATA_ROW = 1;
var HEADER_ROW = 2;
var DATA_START_ROW = 3;

var META_LABEL_COL = 1;  // Column A
var META_ITEM_NUM_COL = 2;  // Column B
var META_ITEM_NAME_COL = 3;  // Column C
var META_ITEM_DESC_COL = 4;  // Column D

/**
 * Creates a new rack configuration tab
 * Prompts user for rack name and parent item
 */
function createNewRackConfiguration() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Step 1: Ask for rack name
  var nameResponse = ui.prompt(
    'New Rack Configuration',
    'Enter a name for this rack configuration (e.g., "Hyperscale Compute Rack"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (nameResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var rackName = nameResponse.getResponseText().trim();
  if (!rackName) {
    ui.alert('Error', 'Rack name cannot be empty.', ui.ButtonSet.OK);
    return;
  }

  // Step 2: Ask if linking to existing item or creating placeholder
  var linkResponse = ui.alert(
    'Parent Item',
    'Do you want to link this rack configuration to an existing Arena item?\n\n' +
    'Yes - Select existing item from Arena\n' +
    'No - Enter placeholder item number (will create in Arena later)',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (linkResponse === ui.Button.CANCEL) {
    return;
  }

  var rackItemNumber, rackItemName, rackItemDescription;

  if (linkResponse === ui.Button.YES) {
    // Option A: Select from Arena using Item Picker
    // For now, use a simple prompt (TODO: enhance with filtered Item Picker)
    var itemResponse = ui.prompt(
      'Select Parent Item',
      'Enter the Arena item number for this rack:',
      ui.ButtonSet.OK_CANCEL
    );

    if (itemResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    rackItemNumber = itemResponse.getResponseText().trim();
    if (!rackItemNumber) {
      ui.alert('Error', 'Item number cannot be empty.', ui.ButtonSet.OK);
      return;
    }

    // Fetch item details from Arena
    try {
      var arenaClient = new ArenaAPIClient();
      var item = arenaClient.getItemByNumber(rackItemNumber);

      if (!item) {
        ui.alert('Error', 'Item "' + rackItemNumber + '" not found in Arena.', ui.ButtonSet.OK);
        return;
      }

      rackItemName = item.name || item.Name || rackName;
      rackItemDescription = item.description || item.Description || '';

    } catch (error) {
      ui.alert('Error', 'Failed to fetch item from Arena: ' + error.message, ui.ButtonSet.OK);
      return;
    }

  } else {
    // Option B: Enter placeholder item number
    var placeholderResponse = ui.prompt(
      'Placeholder Item Number',
      'Enter a placeholder item number for this rack (e.g., "RACK-001"):\n' +
      'This item will be created in Arena when you push the BOM.',
      ui.ButtonSet.OK_CANCEL
    );

    if (placeholderResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    rackItemNumber = placeholderResponse.getResponseText().trim();
    if (!rackItemNumber) {
      ui.alert('Error', 'Item number cannot be empty.', ui.ButtonSet.OK);
      return;
    }

    rackItemName = rackName;
    rackItemDescription = 'Rack configuration for ' + rackName;
  }

  // Step 3: Check for duplicate parent item number
  var existingRacks = getAllRackConfigTabs();
  for (var i = 0; i < existingRacks.length; i++) {
    if (existingRacks[i].itemNumber === rackItemNumber) {
      ui.alert(
        'Duplicate Item',
        'A rack configuration for item "' + rackItemNumber + '" already exists:\n' +
        'Sheet: "' + existingRacks[i].sheetName + '"\n\n' +
        'Please use a different item number or delete the existing rack config.',
        ui.ButtonSet.OK
      );
      return;
    }
  }

  // Step 4: Create the new sheet
  var sheetName = 'Rack - ' + rackItemNumber + ' (' + rackName + ')';
  var newSheet = ss.insertSheet(sheetName);

  // Step 5: Set up metadata row (Row 1)
  newSheet.getRange(METADATA_ROW, META_LABEL_COL).setValue('PARENT_ITEM');
  newSheet.getRange(METADATA_ROW, META_ITEM_NUM_COL).setValue(rackItemNumber);
  newSheet.getRange(METADATA_ROW, META_ITEM_NAME_COL).setValue(rackItemName);
  newSheet.getRange(METADATA_ROW, META_ITEM_DESC_COL).setValue(rackItemDescription);

  // Format metadata row
  var metaRange = newSheet.getRange(METADATA_ROW, 1, 1, 4);
  metaRange.setBackground('#e8f0fe');
  metaRange.setFontWeight('bold');
  metaRange.setFontColor('#1967d2');

  // Step 6: Set up header row (Row 2)
  var itemColumns = getItemColumns();
  var headers = ['Item Number', 'Name', 'Description', 'Category', 'Lifecycle', 'Qty'];

  // Add configured attribute columns
  itemColumns.forEach(function(col) {
    headers.push(col.header || col.attributeName);
  });

  var headerRange = newSheet.getRange(HEADER_ROW, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // Step 7: Freeze rows 1 and 2
  newSheet.setFrozenRows(HEADER_ROW);

  // Step 8: Set column widths
  newSheet.setColumnWidth(1, 120);  // Item Number
  newSheet.setColumnWidth(2, 200);  // Name
  newSheet.setColumnWidth(3, 300);  // Description
  newSheet.setColumnWidth(4, 150);  // Category
  newSheet.setColumnWidth(5, 120);  // Lifecycle
  newSheet.setColumnWidth(6, 60);   // Qty

  // Step 9: Add instructions in Row 3
  newSheet.getRange(DATA_START_ROW, 1).setValue('Use Item Picker to add components →');
  newSheet.getRange(DATA_START_ROW, 1, 1, headers.length).setFontStyle('italic').setFontColor('#666666');

  // Step 10: Set tab color to cascading blue
  var rackIndex = getAllRackConfigTabs().length;  // Get count including this new one
  var blueColor = getCascadingBlueColor(rackIndex - 1);  // Subtract 1 for zero-based index
  newSheet.setTabColor(blueColor);

  // Step 11: Activate the new sheet
  newSheet.activate();

  ui.alert(
    'Rack Configuration Created',
    'Rack configuration sheet created successfully!\n\n' +
    'Sheet: "' + sheetName + '"\n' +
    'Parent Item: ' + rackItemNumber + '\n\n' +
    'Use the Item Picker to add components to this rack.',
    ui.ButtonSet.OK
  );
}

/**
 * Gets metadata from a rack configuration sheet
 * @param {Sheet} sheet - The sheet to read metadata from
 * @return {Object|null} Metadata object or null if not a rack config
 */
function getRackConfigMetadata(sheet) {
  if (!sheet) return null;

  try {
    var label = sheet.getRange(METADATA_ROW, META_LABEL_COL).getValue();

    Logger.log('getRackConfigMetadata: Checking sheet "' + sheet.getName() + '"');
    Logger.log('getRackConfigMetadata: Cell A1 value = "' + label + '"');
    Logger.log('getRackConfigMetadata: Expected value = "PARENT_ITEM"');

    if (label !== 'PARENT_ITEM') {
      Logger.log('getRackConfigMetadata: NOT a rack config sheet (metadata label mismatch)');
      return null;  // Not a rack config sheet
    }

    Logger.log('getRackConfigMetadata: IS a rack config sheet!');

    return {
      itemNumber: sheet.getRange(METADATA_ROW, META_ITEM_NUM_COL).getValue(),
      itemName: sheet.getRange(METADATA_ROW, META_ITEM_NAME_COL).getValue(),
      description: sheet.getRange(METADATA_ROW, META_ITEM_DESC_COL).getValue(),
      sheetName: sheet.getName(),
      sheet: sheet
    };
  } catch (error) {
    Logger.log('Error reading metadata from sheet "' + sheet.getName() + '": ' + error.message);
    return null;
  }
}

/**
 * Checks if a sheet is a rack configuration sheet
 * @param {Sheet} sheet - Sheet to check
 * @return {boolean} True if rack config sheet
 */
function isRackConfigSheet(sheet) {
  return getRackConfigMetadata(sheet) !== null;
}

/**
 * Gets all rack configuration tabs in the spreadsheet
 * @return {Array<Object>} Array of rack config metadata objects
 */
function getAllRackConfigTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var rackConfigs = [];

  sheets.forEach(function(sheet) {
    var metadata = getRackConfigMetadata(sheet);
    if (metadata) {
      // Add child item count
      var lastRow = sheet.getLastRow();
      var childItemCount = Math.max(0, lastRow - HEADER_ROW);
      metadata.childItemCount = childItemCount;

      rackConfigs.push(metadata);
    }
  });

  return rackConfigs;
}

/**
 * Finds a rack configuration tab by parent item number
 * @param {string} itemNumber - Parent rack item number to search for
 * @return {Sheet|null} Rack config sheet or null if not found
 */
function findRackConfigTab(itemNumber) {
  if (!itemNumber) return null;

  var rackConfigs = getAllRackConfigTabs();

  for (var i = 0; i < rackConfigs.length; i++) {
    if (rackConfigs[i].itemNumber === itemNumber) {
      return rackConfigs[i].sheet;
    }
  }

  return null;
}

/**
 * Updates metadata for a rack configuration sheet
 * @param {Sheet} sheet - Rack config sheet to update
 * @param {string} itemNumber - New item number
 * @param {string} itemName - New item name
 * @param {string} description - New description
 */
function updateRackConfigMetadata(sheet, itemNumber, itemName, description) {
  if (!isRackConfigSheet(sheet)) {
    throw new Error('Sheet "' + sheet.getName() + '" is not a rack configuration sheet.');
  }

  // Update metadata cells
  sheet.getRange(METADATA_ROW, META_ITEM_NUM_COL).setValue(itemNumber);
  sheet.getRange(METADATA_ROW, META_ITEM_NAME_COL).setValue(itemName);
  sheet.getRange(METADATA_ROW, META_ITEM_DESC_COL).setValue(description);

  // Update sheet name
  var newSheetName = 'Rack - ' + itemNumber + ' (' + itemName + ')';
  sheet.setName(newSheetName);

  Logger.log('Updated rack config metadata: ' + newSheetName);
}

/**
 * Gets all child items from a rack configuration sheet
 * @param {Sheet} sheet - Rack config sheet
 * @return {Array<Object>} Array of child items with quantities
 */
function getRackConfigChildren(sheet) {
  if (!isRackConfigSheet(sheet)) {
    throw new Error('Sheet "' + sheet.getName() + '" is not a rack configuration sheet.');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    return [];  // No data rows
  }

  var dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 5);  // Get first 5 columns
  var values = dataRange.getValues();

  var children = [];

  values.forEach(function(row) {
    var itemNumber = row[0];
    var itemName = row[1];
    var description = row[2];
    var category = row[3];
    var qty = row[4];

    // Skip empty rows and instruction row
    if (!itemNumber || typeof itemNumber !== 'string') {
      return;
    }

    // Skip instruction text
    if (itemNumber.indexOf('Use Item Picker') !== -1) {
      return;
    }

    // Validate quantity
    if (!qty || isNaN(qty) || qty <= 0) {
      qty = 1;  // Default to 1
    }

    children.push({
      itemNumber: itemNumber,
      name: itemName,
      description: description,
      category: category,
      quantity: Number(qty)
    });
  });

  return children;
}

/**
 * Validates all rack configurations
 * Checks for issues like missing items, invalid quantities, etc.
 * @return {Array<string>} Array of warning messages
 */
function validateRackConfigurations() {
  var warnings = [];
  var rackConfigs = getAllRackConfigTabs();

  rackConfigs.forEach(function(config) {
    // Check if parent item exists in Arena
    try {
      var arenaClient = new ArenaAPIClient();
      var item = arenaClient.getItemByNumber(config.itemNumber);
      if (!item) {
        warnings.push('Rack "' + config.sheetName + '": Parent item "' + config.itemNumber + '" not found in Arena');
      }
    } catch (error) {
      warnings.push('Rack "' + config.sheetName + '": Failed to validate parent item - ' + error.message);
    }

    // Check for child items
    var children = getRackConfigChildren(config.sheet);
    if (children.length === 0) {
      warnings.push('Rack "' + config.sheetName + '": No child items configured');
    }

    // Validate child quantities
    children.forEach(function(child) {
      if (child.quantity <= 0) {
        warnings.push('Rack "' + config.sheetName + '": Item "' + child.itemNumber + '" has invalid quantity: ' + child.quantity);
      }
    });
  });

  return warnings;
}

/**
 * Gets a cascading blue color for rack tabs
 * Returns different shades of blue based on index
 * @param {number} index - Zero-based index of the rack
 * @return {string} Hex color code
 */
function getCascadingBlueColor(index) {
  // Array of blue shades from darker to lighter
  var blueShades = [
    '#0d47a1', // Dark blue
    '#1565c0', // Medium dark blue
    '#1976d2', // Medium blue
    '#1e88e5', // Medium light blue
    '#2196f3', // Light blue
    '#42a5f5', // Lighter blue
    '#64b5f6', // Very light blue
    '#90caf9'  // Lightest blue
  ];

  // Cycle through colors if we have more racks than colors
  return blueShades[index % blueShades.length];
}
