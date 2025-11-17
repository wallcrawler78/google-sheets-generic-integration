/**
 * BOM Builder
 * Handles building, pushing, and pulling Bills of Materials between Google Sheets and Arena
 */

/**
 * Pulls a BOM from Arena and populates the current sheet
 * @param {string} itemNumber - Arena item number to pull BOM for
 * @return {Object} Result object with success status and message
 */
function pullBOM(itemNumber) {
  try {
    var client = new ArenaAPIClient();

    // Find the item by number
    var searchResults = client.searchItems(itemNumber);
    var items = searchResults.results || searchResults.Results || [];

    if (items.length === 0) {
      throw new Error('Item not found: ' + itemNumber);
    }

    var item = items[0];
    var itemGuid = item.guid || item.Guid;

    Logger.log('Found item: ' + itemNumber + ' (GUID: ' + itemGuid + ')');

    // Get the BOM for this item
    var bomData = client.makeRequest('/items/' + itemGuid + '/bom', { method: 'GET' });
    var bomLines = bomData.results || bomData.Results || [];

    Logger.log('Retrieved ' + bomLines.length + ' BOM lines');

    // Get the active sheet
    var sheet = SpreadsheetApp.getActiveSheet();

    // Clear existing data (except headers)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
    }

    // Populate the sheet with BOM data
    populateBOMToSheet(sheet, bomLines, item);

    return {
      success: true,
      message: 'Successfully pulled BOM for ' + itemNumber + '\n' + bomLines.length + ' components imported'
    };

  } catch (error) {
    Logger.log('Error pulling BOM: ' + error.message);
    throw error;
  }
}

/**
 * Populates a sheet with BOM data from Arena
 * @param {Sheet} sheet - The sheet to populate
 * @param {Array} bomLines - Array of BOM line objects from Arena
 * @param {Object} parentItem - The parent item object
 */
function populateBOMToSheet(sheet, bomLines, parentItem) {
  var hierarchy = getBOMHierarchy();
  var columns = getItemColumns();
  var categoryColors = getCategoryColors();

  // Set up headers if not already present
  var headers = ['Level', 'Qty', 'Item Number', 'Item Name', 'Category'];

  // Add configured attribute columns
  columns.forEach(function(col) {
    headers.push(col.header || col.attributeName);
  });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');

  // Process BOM lines
  var rowData = [];

  bomLines.forEach(function(line) {
    var item = line.item || line.Item;
    var quantity = line.quantity || line.Quantity || 1;
    var level = line.level || line.Level || 0;

    var itemNumber = item.number || item.Number || '';
    var itemName = item.name || item.Name || item.description || item.Description || '';
    var categoryName = item.category ? (item.category.name || item.category.Name) : '';

    var row = [
      level,
      quantity,
      itemNumber,
      itemName,
      categoryName
    ];

    // Add configured attributes
    columns.forEach(function(col) {
      var value = getAttributeValue(item, col.attributeGuid);
      row.push(value || '');
    });

    rowData.push(row);
  });

  // Write data to sheet
  if (rowData.length > 0) {
    sheet.getRange(2, 1, rowData.length, headers.length).setValues(rowData);

    // Apply formatting
    formatBOMSheet(sheet, rowData.length + 1);
  }

  // Auto-resize columns
  for (var i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
}

/**
 * Formats a BOM sheet with colors and indentation
 * @param {Sheet} sheet - The sheet to format
 * @param {number} lastRow - The last row with data
 */
function formatBOMSheet(sheet, lastRow) {
  var categoryColors = getCategoryColors();

  // Apply category colors to rows
  for (var row = 2; row <= lastRow; row++) {
    var level = sheet.getRange(row, 1).getValue();
    var category = sheet.getRange(row, 5).getValue();

    // Apply indentation based on level
    var itemNumberCell = sheet.getRange(row, 3);
    var currentValue = itemNumberCell.getValue();
    var indent = '  '.repeat(level);
    itemNumberCell.setValue(indent + currentValue);

    // Apply category color
    var color = getCategoryColor(category);
    if (color) {
      sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(color);
    }
  }

  // Freeze header row
  sheet.setFrozenRows(1);
}

/**
 * Pushes the current sheet BOM to Arena
 * @param {boolean} createNew - If true, creates a new parent item; if false, updates existing
 * @return {Object} Result object with success status and message
 */
function pushBOM() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();

    // Build BOM structure from sheet
    var bomStructure = buildBOMStructure(sheet);

    if (!bomStructure || bomStructure.lines.length === 0) {
      throw new Error('No BOM data found in sheet');
    }

    // Prompt user for parent item
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      'Push BOM to Arena',
      'Enter the parent item number (or leave blank to create new):',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      return { success: false, message: 'Cancelled by user' };
    }

    var parentItemNumber = response.getResponseText().trim();

    var client = new ArenaAPIClient();
    var parentGuid;
    var createNew = !parentItemNumber;

    if (createNew) {
      // Create new parent item
      var newItemResponse = ui.prompt(
        'Create New Item',
        'Enter name for new parent item:',
        ui.ButtonSet.OK_CANCEL
      );

      if (newItemResponse.getSelectedButton() !== ui.Button.OK) {
        return { success: false, message: 'Cancelled by user' };
      }

      var newItemName = newItemResponse.getResponseText().trim();
      if (!newItemName) {
        throw new Error('Item name is required');
      }

      // Create the parent item in Arena
      var newItem = client.createItem({
        name: newItemName,
        category: bomStructure.rootCategory
      });

      parentGuid = newItem.guid || newItem.Guid;
      parentItemNumber = newItem.number || newItem.Number;

      Logger.log('Created new parent item: ' + parentItemNumber);

    } else {
      // Find existing parent item
      var searchResults = client.searchItems(parentItemNumber);
      var items = searchResults.results || searchResults.Results || [];

      if (items.length === 0) {
        throw new Error('Parent item not found: ' + parentItemNumber);
      }

      parentGuid = items[0].guid || items[0].Guid;
      Logger.log('Found parent item: ' + parentItemNumber);
    }

    // Push BOM lines to Arena
    syncBOMToArena(client, parentGuid, bomStructure.lines);

    return {
      success: true,
      message: 'Successfully pushed BOM to ' + parentItemNumber + '\n' +
               bomStructure.lines.length + ' components uploaded'
    };

  } catch (error) {
    Logger.log('Error pushing BOM: ' + error.message);
    throw error;
  }
}

/**
 * Builds a BOM structure from the current sheet layout
 * @param {Sheet} sheet - The sheet containing BOM data
 * @return {Object} BOM structure with lines array
 */
function buildBOMStructure(sheet) {
  var data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return { lines: [], rootCategory: null };
  }

  var headers = data[0];
  var lines = [];
  var rootCategory = null;

  // Find column indices
  var levelCol = headers.indexOf('Level');
  var qtyCol = headers.indexOf('Qty');
  var itemNumberCol = headers.indexOf('Item Number');
  var categoryCol = headers.indexOf('Category');

  if (itemNumberCol === -1) {
    // Try alternate column names
    itemNumberCol = headers.findIndex(function(h) {
      return h.toLowerCase().indexOf('item') !== -1 && h.toLowerCase().indexOf('number') !== -1;
    });
  }

  if (itemNumberCol === -1) {
    throw new Error('Could not find Item Number column in sheet');
  }

  // Process each row (skip header)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var itemNumber = row[itemNumberCol];

    // Skip empty rows
    if (!itemNumber || itemNumber.toString().trim() === '') {
      continue;
    }

    // Remove indentation from item number
    itemNumber = itemNumber.toString().replace(/^\s+/, '');

    var level = levelCol !== -1 ? (row[levelCol] || 0) : 0;
    var quantity = qtyCol !== -1 ? (row[qtyCol] || 1) : 1;
    var category = categoryCol !== -1 ? row[categoryCol] : '';

    // Track root category
    if (level === 0 && !rootCategory) {
      rootCategory = category;
    }

    lines.push({
      level: parseInt(level, 10),
      itemNumber: itemNumber,
      quantity: parseFloat(quantity),
      category: category
    });
  }

  return {
    lines: lines,
    rootCategory: rootCategory
  };
}

/**
 * Syncs BOM lines to Arena
 * @param {ArenaAPIClient} client - Arena API client
 * @param {string} parentGuid - Parent item GUID
 * @param {Array} bomLines - Array of BOM line objects with itemGuid and itemNumber
 */
/**
 * Syncs BOM lines to Arena (creates/updates BOM for a parent item)
 * @param {ArenaAPIClient} client - Arena API client
 * @param {string} parentGuid - Parent item GUID
 * @param {Array} bomLines - Array of BOM line objects {itemNumber, itemGuid, quantity, level, attributes}
 * @param {Object} options - Optional configuration {bomAttributes: {itemNumber: attributeValue}}
 */
function syncBOMToArena(client, parentGuid, bomLines, options) {
  options = options || {};
  var bomAttributes = options.bomAttributes || {};

  // Input validation - fail fast if critical data is missing
  if (!client) {
    throw new Error('syncBOMToArena: ArenaAPIClient is required');
  }
  if (!parentGuid) {
    throw new Error('syncBOMToArena: Parent item GUID is required');
  }
  if (!bomLines || !Array.isArray(bomLines)) {
    throw new Error('syncBOMToArena: bomLines must be an array');
  }

  // Validate all BOM lines have required fields before proceeding
  Logger.log('Validating ' + bomLines.length + ' BOM lines before sync...');
  var validationErrors = [];

  for (var i = 0; i < bomLines.length; i++) {
    var line = bomLines[i];
    var lineRef = 'BOM line ' + (i + 1) + ' (' + (line.itemNumber || 'unknown') + ')';

    if (!line.itemNumber) {
      validationErrors.push(lineRef + ': missing itemNumber');
    }
    if (!line.itemGuid) {
      validationErrors.push(lineRef + ': missing itemGuid');
    }
    if (typeof line.quantity === 'undefined' || line.quantity === null) {
      validationErrors.push(lineRef + ': missing quantity');
    }
    if (typeof line.level === 'undefined' || line.level === null) {
      validationErrors.push(lineRef + ': missing level');
    }
  }

  // If any validation errors, fail immediately with comprehensive error message
  if (validationErrors.length > 0) {
    var errorMsg = 'syncBOMToArena: BOM validation failed with ' + validationErrors.length + ' error(s):\n' +
                   validationErrors.join('\n');
    Logger.log('ERROR: ' + errorMsg);
    throw new Error(errorMsg);
  }

  Logger.log('✓ All BOM lines validated successfully');

  // First, get existing BOM lines for this item
  var existingBOM = [];
  try {
    var bomData = client.makeRequest('/items/' + parentGuid + '/bom', { method: 'GET' });
    existingBOM = bomData.results || bomData.Results || [];
  } catch (error) {
    Logger.log('No existing BOM found (this is OK for new items)');
  }

  // Delete existing BOM lines
  if (existingBOM.length > 0) {
    Logger.log('Deleting ' + existingBOM.length + ' existing BOM lines...');

    existingBOM.forEach(function(line) {
      var lineGuid = line.guid || line.Guid;
      try {
        client.makeRequest('/items/' + parentGuid + '/bom/' + lineGuid, { method: 'DELETE' });
      } catch (deleteError) {
        Logger.log('Error deleting BOM line: ' + deleteError.message);
      }
    });
  }

  // Add new BOM lines
  Logger.log('Adding ' + bomLines.length + ' BOM lines...');

  bomLines.forEach(function(line, index) {
    try {
      // All validation done at start - safe to proceed
      // Create BOM line
      var bomLineData = {
        item: {
          guid: line.itemGuid
        },
        quantity: line.quantity,
        level: line.level,
        lineNumber: index + 1
      };

      // Add custom BOM attributes if provided
      if (bomAttributes[line.itemNumber]) {
        var attrValue = bomAttributes[line.itemNumber];
        // additionalAttributes is how Arena accepts custom BOM-level attributes
        bomLineData.additionalAttributes = attrValue;
        Logger.log('Adding BOM attribute for ' + line.itemNumber + ': ' + JSON.stringify(attrValue));
      }

      client.makeRequest('/items/' + parentGuid + '/bom', {
        method: 'POST',
        payload: bomLineData
      });

      Logger.log('✓ Added BOM line ' + (index + 1) + ': ' + line.itemNumber + ' (GUID: ' + line.itemGuid + ')');

      // Add delay to avoid rate limiting
      Utilities.sleep(100);

    } catch (error) {
      var errorMsg = 'Failed to add BOM line ' + (index + 1) + ' (' + line.itemNumber + '): ' + error.message;

      // Check if this is an attribute error and provide helpful guidance
      if (error.message && (error.message.indexOf('additional attribute') !== -1 ||
                           error.message.indexOf('additionalAttributes') !== -1)) {
        errorMsg += '\n\nThis error suggests the configured BOM attribute may not be valid. ' +
                   'Please reconfigure using: Arena Data Center > Configuration > Rack BOM Location Setting';
      }

      Logger.log('ERROR: ' + errorMsg);
      throw new Error(errorMsg);  // Fail loudly - don't continue with incomplete BOM
    }
  });
}

/**
 * Calculates aggregated quantities across multiple sheets/configurations
 * @param {Array<string>} sheetNames - Array of sheet names to aggregate
 * @return {Object} Map of item numbers to total quantities
 */
function aggregateQuantities(sheetNames) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var quantities = {};

  sheetNames.forEach(function(sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('Warning: Sheet not found: ' + sheetName);
      return;
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    // Find columns
    var itemNumberCol = headers.findIndex(function(h) {
      return h.toLowerCase().indexOf('item') !== -1 && h.toLowerCase().indexOf('number') !== -1;
    });

    var qtyCol = headers.indexOf('Qty');

    if (itemNumberCol === -1) {
      Logger.log('Warning: Item Number column not found in sheet: ' + sheetName);
      return;
    }

    // Process rows
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var itemNumber = row[itemNumberCol];
      var qty = qtyCol !== -1 ? (row[qtyCol] || 1) : 1;

      if (!itemNumber || itemNumber.toString().trim() === '') {
        continue;
      }

      // Remove indentation
      itemNumber = itemNumber.toString().replace(/^\s+/, '').trim();

      if (quantities[itemNumber]) {
        quantities[itemNumber] += parseFloat(qty);
      } else {
        quantities[itemNumber] = parseFloat(qty);
      }
    }
  });

  return quantities;
}

/**
 * Creates a consolidated BOM from multiple rack configurations
 * @param {Array<string>} rackSheetNames - Array of rack sheet names
 * @return {Object} Result with consolidated BOM data
 */
function consolidateBOM(rackSheetNames) {
  try {
    var quantities = aggregateQuantities(rackSheetNames);

    // Get item details from Arena
    var client = new ArenaAPIClient();
    var consolidatedLines = [];

    for (var itemNumber in quantities) {
      try {
        var searchResults = client.searchItems(itemNumber);
        var items = searchResults.results || searchResults.Results || [];

        if (items.length > 0) {
          var item = items[0];
          consolidatedLines.push({
            itemNumber: itemNumber,
            itemName: item.name || item.Name || '',
            category: item.category ? (item.category.name || item.category.Name) : '',
            totalQuantity: quantities[itemNumber]
          });
        } else {
          // Item not found in Arena, still include it
          consolidatedLines.push({
            itemNumber: itemNumber,
            itemName: 'Unknown',
            category: 'Unknown',
            totalQuantity: quantities[itemNumber]
          });
        }

        Utilities.sleep(50); // Avoid rate limiting

      } catch (error) {
        Logger.log('Error fetching item details for ' + itemNumber + ': ' + error.message);
      }
    }

    // Sort by category, then item number
    consolidatedLines.sort(function(a, b) {
      if (a.category !== b.category) {
        return a.category.localeCompare(b.category);
      }
      return a.itemNumber.localeCompare(b.itemNumber);
    });

    return {
      success: true,
      lines: consolidatedLines,
      totalItems: consolidatedLines.length,
      sourceSheets: rackSheetNames
    };

  } catch (error) {
    Logger.log('Error consolidating BOM: ' + error.message);
    throw error;
  }
}

/**
 * Creates a new sheet with consolidated BOM from overview layout
 * Scans overview sheet for rack placements and builds hierarchical BOM
 */
function createConsolidatedBOMSheet() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Find overview sheet (look for sheet with "overview" in name)
  var allSheets = spreadsheet.getSheets();
  var overviewSheet = null;

  for (var i = 0; i < allSheets.length; i++) {
    var sheetName = allSheets[i].getName().toLowerCase();
    if (sheetName.indexOf('overview') !== -1) {
      overviewSheet = allSheets[i];
      break;
    }
  }

  if (!overviewSheet) {
    ui.alert('No Overview Sheet',
      'Could not find an Overview sheet.\n\n' +
      'Please create an overview layout sheet first.',
      ui.ButtonSet.OK);
    return;
  }

  Logger.log('Using overview sheet: ' + overviewSheet.getName());

  // Build consolidated BOM from overview
  try {
    ui.alert('Building Consolidated BOM',
      'Scanning overview layout and rack configurations...\nThis may take a moment.',
      ui.ButtonSet.OK);

    var bomData = buildConsolidatedBOMFromOverview(overviewSheet);

    if (!bomData || bomData.lines.length === 0) {
      ui.alert('No Data',
        'No BOM data could be generated from the overview.\n\n' +
        'Make sure:\n' +
        '1. Overview sheet has rack items placed\n' +
        '2. Rack configuration tabs exist for those racks\n' +
        '3. Rack configs have child items',
        ui.ButtonSet.OK);
      return;
    }

    // Create new sheet
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    var newSheetName = 'Consolidated BOM';
    var existingSheet = spreadsheet.getSheetByName(newSheetName);

    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    var newSheet = spreadsheet.insertSheet(newSheetName);

    // Add summary header
    newSheet.getRange(1, 1).setValue('Consolidated BOM');
    newSheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');

    newSheet.getRange(2, 1).setValue('Generated: ' + timestamp);
    newSheet.getRange(3, 1).setValue('Overview Sheet: ' + overviewSheet.getName());
    newSheet.getRange(4, 1).setValue('Total Items: ' + bomData.totalUniqueItems);
    newSheet.getRange(5, 1).setValue('Total Racks: ' + bomData.totalRacks);

    // Column headers (starting at row 7)
    var headers = ['Level', 'Item Number', 'Name', 'Description', 'Quantity', 'Category'];
    newSheet.getRange(7, 1, 1, headers.length).setValues([headers]);
    newSheet.getRange(7, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');

    // Write BOM data
    var rowData = [];
    bomData.lines.forEach(function(line) {
      rowData.push([
        line.level,
        line.itemNumber,
        line.name,
        line.description,
        line.quantity,
        line.category
      ]);
    });

    if (rowData.length > 0) {
      newSheet.getRange(8, 1, rowData.length, headers.length).setValues(rowData);

      // Apply category colors
      for (var i = 0; i < rowData.length; i++) {
        var category = rowData[i][5];
        var color = getCategoryColor(category);
        if (color) {
          newSheet.getRange(i + 8, 1, 1, headers.length).setBackground(color);
        }
      }

      // Format level column with indentation
      for (var i = 0; i < rowData.length; i++) {
        var level = rowData[i][0];
        var itemNumberCell = newSheet.getRange(i + 8, 2);
        var currentValue = itemNumberCell.getValue();
        var indent = '  '.repeat(level);
        itemNumberCell.setValue(indent + currentValue);
      }
    }

    // Format sheet
    newSheet.setFrozenRows(7);
    newSheet.setColumnWidth(1, 60);   // Level
    newSheet.setColumnWidth(2, 150);  // Item Number
    newSheet.setColumnWidth(3, 250);  // Name
    newSheet.setColumnWidth(4, 300);  // Description
    newSheet.setColumnWidth(5, 80);   // Quantity
    newSheet.setColumnWidth(6, 150);  // Category

    // Add borders
    newSheet.getRange(7, 1, rowData.length + 1, headers.length)
      .setBorder(true, true, true, true, true, true);

    // Set purple tab color to match other system tabs
    newSheet.setTabColor('#9c27b0');

    // Move to end of sheet list
    spreadsheet.setActiveSheet(newSheet);
    spreadsheet.moveActiveSheet(spreadsheet.getNumSheets());

    // Show success message
    ui.alert('Success!',
      'Consolidated BOM created successfully!\n\n' +
      'Total unique items: ' + bomData.totalUniqueItems + '\n' +
      'Total rack instances: ' + bomData.totalRacks + '\n' +
      'BOM lines: ' + bomData.lines.length,
      ui.ButtonSet.OK);

    // Activate the new sheet
    newSheet.activate();

  } catch (error) {
    Logger.log('Error creating consolidated BOM: ' + error.message + '\n' + error.stack);
    ui.alert('Error',
      'Failed to create consolidated BOM:\n\n' + error.message,
      ui.ButtonSet.OK);
  }
}

/**
 * Builds consolidated BOM data from overview sheet
 * Scans for rack placements, looks up rack configs, multiplies quantities
 * @param {Sheet} overviewSheet - The overview/layout sheet
 * @return {Object} BOM data with lines array and stats
 */
function buildConsolidatedBOMFromOverview(overviewSheet) {
  var hierarchy = getBOMHierarchy();
  var arenaClient = new ArenaAPIClient();

  // Step 1: Scan overview sheet for all rack placements
  Logger.log('Step 1: Scanning overview sheet for rack placements...');
  var rackPlacements = scanOverviewForRacks(overviewSheet);

  Logger.log('Found ' + Object.keys(rackPlacements).length + ' unique rack types');
  Logger.log('Total rack instances: ' + getTotalRackCount(rackPlacements));

  if (Object.keys(rackPlacements).length === 0) {
    throw new Error('No rack items found in overview sheet');
  }

  // Step 2: For each unique rack, get its configuration and children
  Logger.log('Step 2: Loading rack configurations and children...');
  var consolidatedItems = {};  // Map: itemNumber => {item data, total quantity}

  for (var rackItemNumber in rackPlacements) {
    var rackCount = rackPlacements[rackItemNumber];
    Logger.log('Processing rack: ' + rackItemNumber + ' (count: ' + rackCount + ')');

    // Find rack config tab
    var rackConfigSheet = findRackConfigTab(rackItemNumber);

    if (!rackConfigSheet) {
      Logger.log('WARNING: No rack config found for ' + rackItemNumber + ', skipping');
      continue;
    }

    // Get all children from rack config
    var children = getRackConfigChildren(rackConfigSheet);
    Logger.log('  Found ' + children.length + ' children in rack config');

    // Add rack itself to consolidated BOM
    var rackMetadata = getRackConfigMetadata(rackConfigSheet);
    if (!consolidatedItems[rackItemNumber]) {
      consolidatedItems[rackItemNumber] = {
        itemNumber: rackItemNumber,
        name: rackMetadata.itemName,
        description: rackMetadata.description,
        category: '', // Will fetch from Arena
        quantity: 0,
        level: 2  // Rack level (will be adjusted based on hierarchy config)
      };
    }
    consolidatedItems[rackItemNumber].quantity += rackCount;

    // Add children with multiplied quantities
    children.forEach(function(child) {
      var totalChildQty = child.quantity * rackCount;

      if (!consolidatedItems[child.itemNumber]) {
        consolidatedItems[child.itemNumber] = {
          itemNumber: child.itemNumber,
          name: child.name,
          description: child.description,
          category: child.category,
          quantity: 0,
          level: 3  // Child level (will be adjusted)
        };
      }

      consolidatedItems[child.itemNumber].quantity += totalChildQty;
    });
  }

  // Step 3: Fetch additional details from Arena and apply hierarchy levels
  Logger.log('Step 3: Fetching item details from Arena and applying hierarchy...');
  var bomLines = [];

  for (var itemNumber in consolidatedItems) {
    var item = consolidatedItems[itemNumber];

    // Fetch from Arena if category not set
    if (!item.category) {
      try {
        var arenaItem = arenaClient.getItemByNumber(itemNumber);
        if (arenaItem) {
          var categoryObj = arenaItem.category || arenaItem.Category || {};
          item.category = categoryObj.name || categoryObj.Name || '';
          if (!item.name) item.name = arenaItem.name || arenaItem.Name || '';
          if (!item.description) item.description = arenaItem.description || arenaItem.Description || '';
        }
      } catch (error) {
        Logger.log('Could not fetch details for ' + itemNumber + ': ' + error.message);
      }
    }

    // Determine BOM level based on category → level mapping
    var bomLevel = getBOMLevelForCategory(item.category, hierarchy);
    if (bomLevel !== null) {
      item.level = bomLevel;
    }

    bomLines.push(item);
  }

  // Step 4: Sort by level, then category, then item number
  bomLines.sort(function(a, b) {
    if (a.level !== b.level) return a.level - b.level;
    if (a.category !== b.category) return a.category.localeCompare(b.category);
    return a.itemNumber.localeCompare(b.itemNumber);
  });

  Logger.log('Built consolidated BOM with ' + bomLines.length + ' unique items');

  return {
    lines: bomLines,
    totalUniqueItems: bomLines.length,
    totalRacks: getTotalRackCount(rackPlacements),
    sourceSheet: overviewSheet.getName()
  };
}

/**
 * Scans overview sheet for rack item placements
 * @param {Sheet} sheet - Overview sheet
 * @return {Object} Map of rack item numbers to instance counts
 */
function scanOverviewForRacks(sheet) {
  var data = sheet.getDataRange().getValues();
  var rackPlacements = {};

  // Scan all cells for rack item numbers
  for (var row = 0; row < data.length; row++) {
    for (var col = 0; col < data[row].length; col++) {
      var cellValue = data[row][col];

      if (!cellValue || typeof cellValue !== 'string') continue;

      var itemNumber = cellValue.trim();

      // Skip empty cells and headers
      if (!itemNumber || itemNumber.length === 0) continue;

      // Check if this is a rack item (has a rack config tab)
      if (findRackConfigTab(itemNumber)) {
        if (rackPlacements[itemNumber]) {
          rackPlacements[itemNumber]++;
        } else {
          rackPlacements[itemNumber] = 1;
        }
      }
    }
  }

  return rackPlacements;
}

/**
 * Gets total count of all rack instances
 * @param {Object} rackPlacements - Map of rack items to counts
 * @return {number} Total count
 */
function getTotalRackCount(rackPlacements) {
  var total = 0;
  for (var rackItem in rackPlacements) {
    total += rackPlacements[rackItem];
  }
  return total;
}

/**
 * Determines BOM level for a given category based on hierarchy configuration
 * @param {string} categoryName - Category name
 * @param {Array} hierarchy - BOM hierarchy configuration
 * @return {number|null} Level number or null if not found
 */
function getBOMLevelForCategory(categoryName, hierarchy) {
  if (!categoryName || !hierarchy) return null;

  for (var i = 0; i < hierarchy.length; i++) {
    if (hierarchy[i].category === categoryName || hierarchy[i].name === categoryName) {
      return hierarchy[i].level;
    }
  }

  return null;
}

/**
 * Pushes the consolidated BOM to Arena
 * Allows user to update existing item or create new
 */
function pushConsolidatedBOMToArena() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Find the consolidated BOM sheet
  var bomSheet = spreadsheet.getSheetByName('Consolidated BOM');

  if (!bomSheet) {
    ui.alert('No Consolidated BOM',
      'Could not find a "Consolidated BOM" sheet.\n\n' +
      'Please create a consolidated BOM first using:\n' +
      'BOM Operations → Create Consolidated BOM',
      ui.ButtonSet.OK);
    return;
  }

  // Ask user: Update existing or create new?
  var response = ui.alert('Push BOM to Arena',
    'Do you want to UPDATE an existing item or CREATE a new item?\n\n' +
    'Yes = Update existing item\n' +
    'No = Create new item',
    ui.ButtonSet.YES_NO_CANCEL);

  if (response === ui.Button.CANCEL) {
    return;
  }

  var updateExisting = (response === ui.Button.YES);
  var parentItemNumber;
  var parentGuid;
  var client = new ArenaAPIClient();

  try {
    if (updateExisting) {
      // Prompt for existing item number
      var itemResponse = ui.prompt('Update Existing Item',
        'Enter the Arena item number to update:',
        ui.ButtonSet.OK_CANCEL);

      if (itemResponse.getSelectedButton() !== ui.Button.OK) {
        return;
      }

      parentItemNumber = itemResponse.getResponseText().trim();

      if (!parentItemNumber) {
        ui.alert('Error', 'Item number cannot be empty.', ui.ButtonSet.OK);
        return;
      }

      // Find the item in Arena
      var item = client.getItemByNumber(parentItemNumber);

      if (!item) {
        ui.alert('Item Not Found',
          'Item "' + parentItemNumber + '" not found in Arena.\n\n' +
          'Please check the item number and try again.',
          ui.ButtonSet.OK);
        return;
      }

      parentGuid = item.guid || item.Guid;
      Logger.log('Found existing item: ' + parentItemNumber + ' (' + parentGuid + ')');

    } else {
      // Create new item
      var newItemResponse = ui.prompt('Create New Item',
        'Enter details for the new parent item:\n\n' +
        'Item Number (or leave blank for auto-generation):',
        ui.ButtonSet.OK_CANCEL);

      if (newItemResponse.getSelectedButton() !== ui.Button.OK) {
        return;
      }

      var newItemNumber = newItemResponse.getResponseText().trim();

      var nameResponse = ui.prompt('Create New Item',
        'Enter the item name:',
        ui.ButtonSet.OK_CANCEL);

      if (nameResponse.getSelectedButton() !== ui.Button.OK) {
        return;
      }

      var newItemName = nameResponse.getResponseText().trim();

      if (!newItemName) {
        ui.alert('Error', 'Item name is required.', ui.ButtonSet.OK);
        return;
      }

      // Get category for new item (use first level from hierarchy or ask)
      var hierarchy = getBOMHierarchy();
      var defaultCategory = hierarchy.length > 0 ? hierarchy[0].category : '';

      // Create the item
      Logger.log('Creating new item in Arena...');
      var newItemData = {
        name: newItemName
      };

      if (newItemNumber) {
        newItemData.number = newItemNumber;
      }

      if (defaultCategory) {
        newItemData.category = defaultCategory;
      }

      var createdItem = client.createItem(newItemData);

      parentGuid = createdItem.guid || createdItem.Guid;
      parentItemNumber = createdItem.number || createdItem.Number;

      Logger.log('Created new item: ' + parentItemNumber + ' (' + parentGuid + ')');
    }

    // Read BOM data from sheet
    Logger.log('Reading BOM data from sheet...');
    var bomLines = readBOMFromSheet(bomSheet);

    if (bomLines.length === 0) {
      ui.alert('No Data',
        'No BOM lines found in the consolidated BOM sheet.',
        ui.ButtonSet.OK);
      return;
    }

    // Confirm before pushing
    var confirmResponse = ui.alert('Confirm Push',
      'Ready to push BOM to Arena:\n\n' +
      'Parent Item: ' + parentItemNumber + '\n' +
      'BOM Lines: ' + bomLines.length + '\n\n' +
      (updateExisting ? 'This will DELETE the existing BOM and replace it.\n\n' : '') +
      'Continue?',
      ui.ButtonSet.YES_NO);

    if (confirmResponse !== ui.Button.YES) {
      ui.alert('Cancelled', 'BOM push cancelled.', ui.ButtonSet.OK);
      return;
    }

    // Push to Arena
    Logger.log('Pushing BOM to Arena...');
    syncBOMToArena(client, parentGuid, bomLines);

    ui.alert('Success!',
      'BOM pushed successfully to Arena!\n\n' +
      'Parent Item: ' + parentItemNumber + '\n' +
      'BOM Lines: ' + bomLines.length + '\n\n' +
      'You can view the BOM in Arena PLM.',
      ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('Error pushing BOM to Arena: ' + error.message + '\n' + error.stack);
    ui.alert('Error',
      'Failed to push BOM to Arena:\n\n' + error.message,
      ui.ButtonSet.OK);
  }
}

/**
 * Reads BOM data from a consolidated BOM sheet
 * @param {Sheet} sheet - Consolidated BOM sheet
 * @return {Array} Array of BOM line objects
 */
function readBOMFromSheet(sheet) {
  var data = sheet.getDataRange().getValues();
  var bomLines = [];

  // Find header row (look for row with "Level", "Item Number", etc.)
  var headerRow = -1;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[0] === 'Level' || row[1] === 'Item Number') {
      headerRow = i;
      break;
    }
  }

  if (headerRow === -1) {
    throw new Error('Could not find header row in BOM sheet');
  }

  var headers = data[headerRow];
  var levelCol = headers.indexOf('Level');
  var itemNumberCol = headers.indexOf('Item Number');
  var qtyCol = headers.indexOf('Quantity');

  if (itemNumberCol === -1) {
    throw new Error('Could not find "Item Number" column');
  }

  // Read data rows (skip header and any summary rows before it)
  for (var i = headerRow + 1; i < data.length; i++) {
    var row = data[i];
    var itemNumber = row[itemNumberCol];

    // Skip empty rows
    if (!itemNumber || typeof itemNumber !== 'string') continue;

    // Remove indentation
    itemNumber = itemNumber.toString().replace(/^\s+/, '').trim();

    if (!itemNumber) continue;

    var level = levelCol !== -1 ? (row[levelCol] || 0) : 0;
    var quantity = qtyCol !== -1 ? (row[qtyCol] || 1) : 1;

    bomLines.push({
      level: parseInt(level, 10),
      itemNumber: itemNumber,
      quantity: parseFloat(quantity)
    });
  }

  Logger.log('Read ' + bomLines.length + ' BOM lines from sheet');
  return bomLines;
}


/**
 * Identifies custom racks that need Arena items created
 * @param {Array} rackItemNumbers - Array of rack item numbers from overview
 * @return {Array} Array of custom rack objects {itemNumber, metadata, sheet}
 */
function identifyCustomRacks(rackItemNumbers) {
  var client = new ArenaAPIClient();
  var customRacks = [];

  rackItemNumbers.forEach(function(itemNumber) {
    try {
      // Find the rack config tab
      var rackSheet = findRackConfigTab(itemNumber);
      if (!rackSheet) {
        Logger.log('⚠ No rack config found for: ' + itemNumber + ' - cannot create without config sheet');
        return;
      }

      var metadata = getRackConfigMetadata(rackSheet);

      // CRITICAL: Check Arena FIRST to determine if this is a placeholder
      // A rack can have children locally (BOM populated) but not exist in Arena yet
      Logger.log('Checking if rack ' + itemNumber + ' exists in Arena...');
      var arenaItem = client.getItemByNumber(itemNumber);

      if (!arenaItem) {
        // Item doesn't exist in Arena - it's a placeholder rack that needs creation
        var children = getRackConfigChildren(rackSheet);

        if (children && children.length > 0) {
          Logger.log('✓ Custom rack identified (placeholder with BOM): ' + itemNumber + ' (' + children.length + ' children)');
          customRacks.push({
            itemNumber: itemNumber,
            metadata: metadata,
            sheet: rackSheet,
            children: children,
            reason: 'placeholder_with_bom'
          });
        } else {
          Logger.log('✓ Custom rack identified (placeholder without BOM): ' + itemNumber);
          customRacks.push({
            itemNumber: itemNumber,
            metadata: metadata,
            sheet: rackSheet,
            children: [],
            reason: 'placeholder_no_bom'
          });
        }
        return;
      }

      // Item exists in Arena - check if it needs BOM update
      Logger.log('Rack ' + itemNumber + ' found in Arena');
      var itemGuid = arenaItem.guid || arenaItem.Guid;
      var bomData = client.makeRequest('/items/' + itemGuid + '/bom', { method: 'GET' });
      var bomLines = bomData.results || bomData.Results || [];

      // Get local BOM
      var children = getRackConfigChildren(rackSheet);

      if (bomLines.length === 0 && children && children.length > 0) {
        // Item exists but has no BOM in Arena, but has BOM locally - needs BOM sync
        Logger.log('✓ Rack ' + itemNumber + ' exists in Arena but missing BOM (' + children.length + ' children to sync)');
        customRacks.push({
          itemNumber: itemNumber,
          metadata: metadata,
          sheet: rackSheet,
          arenaItem: arenaItem,
          children: children,
          reason: 'needs_bom_sync'
        });
      } else if (bomLines.length > 0 && (!children || children.length === 0)) {
        // Arena has BOM but local config is empty - suggest pull
        Logger.log('ℹ Rack ' + itemNumber + ' exists in Arena with BOM, but local config is empty - consider pulling BOM from Arena');
      } else {
        // Both have BOMs or both are empty - assume synced
        Logger.log('✓ Rack ' + itemNumber + ' appears synchronized (Arena BOM: ' + bomLines.length + ', Local: ' + (children ? children.length : 0) + ')');
      }

    } catch (error) {
      Logger.log('⚠ Error checking rack ' + itemNumber + ': ' + error.message);

      // On error, check if it's a 404 (item not found) - treat as placeholder
      if (error.message && error.message.indexOf('404') !== -1) {
        Logger.log('✓ Rack ' + itemNumber + ' not found in Arena (404) - treating as placeholder');
        var rackSheet = findRackConfigTab(itemNumber);
        if (rackSheet) {
          var children = getRackConfigChildren(rackSheet);
          customRacks.push({
            itemNumber: itemNumber,
            metadata: getRackConfigMetadata(rackSheet),
            sheet: rackSheet,
            children: children || [],
            reason: children && children.length > 0 ? 'placeholder_with_bom' : 'placeholder_no_bom'
          });
        }
      } else {
        // Other errors - be conservative and skip
        Logger.log('⚠ Cannot determine status for ' + itemNumber + ' - skipping to avoid errors');
      }
    }
  });

  return customRacks;
}

/**
 * Creates Arena items for custom racks with user prompts
 * @param {Array} customRacks - Array of custom rack objects
 * @return {Object} Result with created items
 */
function createCustomRackItems(customRacks) {
  if (customRacks.length === 0) {
    return { success: true, createdItems: [] };
  }

  var ui = SpreadsheetApp.getUi();
  var client = new ArenaAPIClient();
  var createdItems = [];

  for (var i = 0; i < customRacks.length; i++) {
    var rack = customRacks[i];

    // Check if Arena item exists but just needs BOM
    if (rack.arenaItem && rack.reason === 'needs_bom_sync') {
      Logger.log('Rack exists in Arena, will add BOM: ' + rack.itemNumber);

      // Get rack children and push BOM
      var children = getRackConfigChildren(rack.sheet);
      var itemGuid = rack.arenaItem.guid || rack.arenaItem.Guid;

      // Convert children to BOM format with GUID lookups
      var bomLines = [];
      for (var j = 0; j < children.length; j++) {
        var child = children[j];
        try {
          Logger.log('Looking up GUID for child component: ' + child.itemNumber);
          var childItem = client.getItemByNumber(child.itemNumber);

          if (!childItem) {
            Logger.log('ERROR: Child component not found in Arena: ' + child.itemNumber);
            throw new Error('Child component not found in Arena: ' + child.itemNumber +
                          '. Needed for rack: ' + rack.itemNumber +
                          '. Please ensure all components exist in Arena before creating rack BOMs.');
          }

          var childGuid = childItem.guid || childItem.Guid;

          bomLines.push({
            itemNumber: child.itemNumber,
            itemGuid: childGuid,  // ✓ Include GUID from Arena lookup
            quantity: child.quantity || 1,
            level: 0
          });

          Logger.log('✓ Found child component GUID: ' + child.itemNumber + ' → ' + childGuid);
        } catch (childError) {
          Logger.log('ERROR looking up child component ' + child.itemNumber + ': ' + childError.message);
          throw childError;  // Fail loudly - don't create incomplete BOMs
        }
      }

      syncBOMToArena(client, itemGuid, bomLines);

      // Update rack status to SYNCED now that BOM has been synced to Arena
      Logger.log('Updating rack status to SYNCED after BOM sync: ' + rack.itemNumber);
      var eventDetails = {
        changesSummary: 'BOM synced to Arena (' + bomLines.length + ' items)',
        details: 'POD push updated existing rack BOM in Arena'
      };
      updateRackSheetStatus(rack.sheet, RACK_STATUS.SYNCED, itemGuid, eventDetails);

      // Log POD push event
      addRackHistoryEvent(rack.itemNumber, HISTORY_EVENT.POD_PUSH, {
        changesSummary: 'Rack BOM updated in Arena during POD push',
        details: bomLines.length + ' BOM items synced to Arena',
        statusAfter: RACK_STATUS.SYNCED
      });

      createdItems.push({
        itemNumber: rack.itemNumber,
        guid: itemGuid,
        updated: true
      });

      continue;
    }

    // Prompt user for rack details
    var promptMsg = '========================================\n' +
                    'CREATING CUSTOM RACK ITEM ' + (i + 1) + ' of ' + customRacks.length + '\n' +
                    '========================================\n\n' +
                    'Rack Item Number: ' + rack.itemNumber + '\n' +
                    'Current Name: ' + (rack.metadata.itemName || 'Not set') + '\n\n' +
                    '----------------------------------------\n' +
                    'Enter a name for this rack in Arena:';

    var nameResponse = ui.prompt('Create Custom Rack (' + (i + 1) + ' of ' + customRacks.length + ')', promptMsg, ui.ButtonSet.OK_CANCEL);

    if (nameResponse.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Error', 'Custom rack creation cancelled. Cannot proceed with POD creation.', ui.ButtonSet.OK);
      return { success: false, message: 'Cancelled by user' };
    }

    var rackName = nameResponse.getResponseText().trim();
    if (!rackName) {
      ui.alert('Error', 'Rack name is required.', ui.ButtonSet.OK);
      return { success: false, message: 'Invalid rack name' };
    }

    // Use clickable category selector dialog
    Logger.log('Prompting user to select category for rack: ' + rack.itemNumber);

    var categorySelection = showCategorySelector(
      'Select Category for Rack ' + rack.itemNumber,
      'Choose the Arena category for this custom rack item (' + (i + 1) + ' of ' + customRacks.length + ')'
    );

    if (!categorySelection) {
      ui.alert('Error', 'Category selection cancelled. Cannot proceed with POD creation.', ui.ButtonSet.OK);
      return { success: false, message: 'Cancelled by user' };
    }

    var selectedCategoryName = categorySelection.name;
    Logger.log('Selected category: ' + selectedCategoryName + ' (GUID: ' + categorySelection.guid + ')');

    // Prompt for description
    var descResponse = ui.prompt(
      'Description for Rack ' + rack.itemNumber,
      'RACK: ' + rack.itemNumber + '\n\n' +
      'Enter a description for this rack in Arena:',
      ui.ButtonSet.OK_CANCEL
    );

    if (descResponse.getSelectedButton() !== ui.Button.OK) {
      return { success: false, message: 'Cancelled by user' };
    }

    var description = descResponse.getResponseText().trim();

    // Create the item in Arena
    try {
      Logger.log('Creating rack item: ' + rack.itemNumber + ' with category: ' + selectedCategoryName + ' (GUID: ' + categorySelection.guid + ')');

      // Arena API expects category as an object with guid, not a simple string
      // NOTE: Don't include 'number' in initial creation - Arena may have auto-numbering enabled
      var newItem = client.createItem({
        name: rackName,
        category: {
          guid: categorySelection.guid
        },
        description: description
      });

      var newItemGuid = newItem.guid || newItem.Guid;
      var newItemNumber = newItem.number || newItem.Number;

      Logger.log('Created rack item in Arena (GUID: ' + newItemGuid + ', auto-number: ' + newItemNumber + ')');

      // Always update with our desired rack number (handles both auto-numbering and manual numbering)
      Logger.log('Updating rack item number to: ' + rack.itemNumber);
      client.updateItem(newItemGuid, { number: rack.itemNumber });
      newItemNumber = rack.itemNumber;
      Logger.log('✓ Rack item number set to: ' + newItemNumber);

      // Get rack children and push BOM
      var children = getRackConfigChildren(rack.sheet);

      // Convert children to BOM format with GUID lookups
      var bomLines = [];
      for (var k = 0; k < children.length; k++) {
        var child = children[k];
        try {
          Logger.log('Looking up GUID for child component: ' + child.itemNumber);
          var childItem = client.getItemByNumber(child.itemNumber);

          if (!childItem) {
            Logger.log('ERROR: Child component not found in Arena: ' + child.itemNumber);
            throw new Error('Child component not found in Arena: ' + child.itemNumber +
                          '. Needed for rack: ' + rack.itemNumber +
                          '. Please ensure all components exist in Arena before creating rack BOMs.');
          }

          var childGuid = childItem.guid || childItem.Guid;

          bomLines.push({
            itemNumber: child.itemNumber,
            itemGuid: childGuid,  // ✓ Include GUID from Arena lookup
            quantity: child.quantity || 1,
            level: 0
          });

          Logger.log('✓ Found child component GUID: ' + child.itemNumber + ' → ' + childGuid);
        } catch (childError) {
          Logger.log('ERROR looking up child component ' + child.itemNumber + ': ' + childError.message);
          throw childError;  // Fail loudly - don't create incomplete BOMs
        }
      }

      syncBOMToArena(client, newItemGuid, bomLines);

      // Update rack config metadata with Arena info
      rack.sheet.getRange(1, 2).setValue(rack.itemNumber);
      rack.sheet.getRange(1, 3).setValue(rackName);
      rack.sheet.getRange(1, 4).setValue(description);

      // Update rack status to SYNCED now that it has been created in Arena with BOM
      Logger.log('Updating rack status to SYNCED after creation: ' + rack.itemNumber + ' (GUID: ' + newItemGuid + ')');
      var eventDetails = {
        changesSummary: 'Rack created in Arena with BOM',
        details: 'POD push created new rack with ' + bomLines.length + ' BOM items'
      };
      updateRackSheetStatus(rack.sheet, RACK_STATUS.SYNCED, newItemGuid, eventDetails);

      // Log POD push event (creation)
      addRackHistoryEvent(rack.itemNumber, HISTORY_EVENT.POD_PUSH, {
        changesSummary: 'Rack created in Arena during POD push',
        details: 'New rack item created with ' + bomLines.length + ' BOM items',
        statusAfter: RACK_STATUS.SYNCED
      });

      createdItems.push({
        itemNumber: rack.itemNumber,
        guid: newItemGuid,
        name: rackName,
        created: true
      });

    } catch (error) {
      Logger.log('Error creating rack item: ' + error.message);
      ui.alert('Error', 'Failed to create rack item: ' + error.message, ui.ButtonSet.OK);
      return { success: false, message: error.message };
    }
  }

  return {
    success: true,
    createdItems: createdItems
  };
}

/**
 * Scans overview sheet row by row for rack placements
 * @param {Sheet} sheet - Overview sheet
 * @return {Array} Array of row objects {rowNumber, positions: [{col, itemNumber, rackCount}]}
 */
function scanOverviewByRow(sheet) {
  var data = sheet.getDataRange().getValues();
  var rowData = [];

  // Find header row (contains "Pos 1", "Pos 2", etc.)
  var headerRow = -1;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    for (var j = 0; j < row.length; j++) {
      if (row[j] && row[j].toString().toLowerCase().indexOf('pos') === 0) {
        headerRow = i;
        break;
      }
    }
    if (headerRow !== -1) break;
  }

  if (headerRow === -1) {
    throw new Error('Could not find position headers in overview sheet');
  }

  var headers = data[headerRow];
  var firstPosCol = -1;

  // Find first position column
  for (var j = 0; j < headers.length; j++) {
    if (headers[j] && headers[j].toString().toLowerCase().indexOf('pos') === 0) {
      firstPosCol = j;
      break;
    }
  }

  // Scan each row after headers
  for (var i = headerRow + 1; i < data.length; i++) {
    var row = data[i];
    var positions = [];
    var rowHasData = false;

    // Scan position columns
    for (var j = firstPosCol; j < row.length; j++) {
      var cellValue = row[j];

      if (!cellValue) continue;

      // Check if it's a rack (try to find config tab)
      var rackSheet = findRackConfigTab(cellValue.toString());
      if (rackSheet) {
        rowHasData = true;
        positions.push({
          col: j,
          positionName: headers[j],
          itemNumber: cellValue.toString()
        });
      }
    }

    if (rowHasData) {
      // Get row number from first column
      var rowNumber = row[0] || (rowData.length + 1);

      rowData.push({
        rowNumber: rowNumber,
        sheetRow: i + 1, // Sheet row index (1-based)
        positions: positions
      });
    }
  }

  Logger.log('Scanned overview: found ' + rowData.length + ' rows with racks');
  return rowData;
}

/**
 * Creates Row items in Arena with BOM position tracking
 * @param {Array} rowData - Array of row objects from scanOverviewByRow
 * @param {string} rowCategory - Category to use for Row items
 * @return {Array} Array of created row items with metadata
 */
function createRowItems(rowData, rowCategory) {
  var ui = SpreadsheetApp.getUi();
  var client = new ArenaAPIClient();
  var rowItems = [];

  for (var i = 0; i < rowData.length; i++) {
    var row = rowData[i];

    // Prompt user for row name
    var promptMsg = '========================================\n' +
                    'CREATING ROW ITEM ' + (i + 1) + ' of ' + rowData.length + '\n' +
                    '========================================\n\n' +
                    'Overview Row Number: ' + row.rowNumber + '\n\n' +
                    'This row contains racks in the following positions:\n';

    row.positions.forEach(function(pos) {
      promptMsg += '  • ' + pos.positionName + ': ' + pos.itemNumber + '\n';
    });

    promptMsg += '\n----------------------------------------\n';
    promptMsg += 'Enter a name for this Row item in Arena:';

    var nameResponse = ui.prompt('Create Row Item (' + (i + 1) + ' of ' + rowData.length + ')', promptMsg, ui.ButtonSet.OK_CANCEL);

    if (nameResponse.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Error', 'Row creation cancelled.', ui.ButtonSet.OK);
      return null;
    }

    var rowName = nameResponse.getResponseText().trim();
    if (!rowName) {
      rowName = 'Row ' + row.rowNumber;
    }

    // Build position names for item description (comma-separated)
    var positionNames = row.positions.map(function(pos) {
      return pos.positionName;
    }).join(', ');

    // Aggregate rack quantities for this row
    var rackCounts = {};
    row.positions.forEach(function(pos) {
      if (!rackCounts[pos.itemNumber]) {
        rackCounts[pos.itemNumber] = 0;
      }
      rackCounts[pos.itemNumber]++;
    });

    // Create row item in Arena
    try {
      Logger.log('=== CREATING ROW ITEM ' + (i + 1) + ' ===');
      Logger.log('Row name: ' + rowName);
      Logger.log('Row category (should be GUID): ' + rowCategory);
      Logger.log('Row category type: ' + typeof rowCategory);
      Logger.log('Row description: Row ' + row.rowNumber + ' with racks in positions: ' + positionNames);

      // Arena API expects category as an object with guid, not a simple string
      var createItemPayload = {
        name: rowName,
        category: {
          guid: rowCategory
        },
        description: 'Row ' + row.rowNumber + ' with racks in positions: ' + positionNames
      };

      Logger.log('Full createItem payload: ' + JSON.stringify(createItemPayload));

      var rowItem = client.createItem(createItemPayload);

      Logger.log('Full response from createItem: ' + JSON.stringify(rowItem));

      var rowItemGuid = rowItem.guid || rowItem.Guid;
      var rowItemNumber = rowItem.number || rowItem.Number;

      Logger.log('✓ Created row item: ' + rowItemNumber + ' (GUID: ' + rowItemGuid + ')');

      // Handle manual item numbering if needed
      if (!rowItemNumber) {
        Logger.log('⚠ Item number is null - category may not have auto-numbering enabled');

        var numberPrompt = '========================================\n' +
                           'ITEM NUMBER REQUIRED\n' +
                           '========================================\n\n' +
                           'The ROW category does not have auto-numbering enabled.\n' +
                           'Please enter an item number for this row:\n\n' +
                           'Row: ' + rowName + '\n' +
                           'GUID: ' + rowItemGuid;

        var numberResponse = ui.prompt('Enter Item Number', numberPrompt, ui.ButtonSet.OK_CANCEL);

        if (numberResponse.getSelectedButton() !== ui.Button.OK) {
          throw new Error('Item number is required but was not provided');
        }

        rowItemNumber = numberResponse.getResponseText().trim();
        if (!rowItemNumber) {
          throw new Error('Item number cannot be empty');
        }

        // Update the item with the user-provided number
        Logger.log('Updating item with manual number: ' + rowItemNumber);
        client.updateItem(rowItemGuid, { number: rowItemNumber });
        Logger.log('✓ Item number set to: ' + rowItemNumber);
      }

      // Position tracking is now handled via BOM attributes (see Rack BOM Location Setting)
      // Each rack on the BOM will have its positions tagged via the configured BOM attribute

      // Create BOM for row (add each rack with its quantity)
      // First, look up GUIDs for each rack from Arena
      var bomLines = [];
      for (var rackNumber in rackCounts) {
        try {
          Logger.log('Looking up GUID for rack: ' + rackNumber);
          var rackItem = client.getItemByNumber(rackNumber);

          if (!rackItem) {
            Logger.log('ERROR: Rack item not found in Arena: ' + rackNumber);
            throw new Error('Rack item not found in Arena: ' + rackNumber + '. Please ensure all rack items exist in Arena before creating rows.');
          }

          var rackGuid = rackItem.guid || rackItem.Guid;

          bomLines.push({
            itemNumber: rackNumber,
            itemGuid: rackGuid,  // Add GUID from Arena lookup
            quantity: rackCounts[rackNumber],
            level: 0
          });

          Logger.log('✓ Found rack GUID: ' + rackNumber + ' → ' + rackGuid);
        } catch (rackError) {
          Logger.log('ERROR looking up rack ' + rackNumber + ': ' + rackError.message);
          throw rackError;  // Stop processing if we can't find a rack
        }
      }

      // Build position mapping for BOM attributes (if configured)
      var bomOptions = {};
      var positionConfig = getBOMPositionAttributeConfig();

      if (positionConfig) {
        Logger.log('✓ Position tracking enabled - using attribute: ' + positionConfig.name);

        // Build rack-to-positions mapping
        var rackPositions = {}; // Map: rackNumber => [positionNames]

        row.positions.forEach(function(pos) {
          if (!rackPositions[pos.itemNumber]) {
            rackPositions[pos.itemNumber] = [];
          }
          rackPositions[pos.itemNumber].push(pos.positionName);
        });

        // Format position values and build BOM attributes map
        var bomAttributes = {};

        for (var rackNumber in rackPositions) {
          var positions = rackPositions[rackNumber];
          var formattedPositions = positions.join(', '); // e.g., "Pos 1, Pos 3, Pos 8"

          // Build additionalAttributes structure for Arena API (must be array format)
          bomAttributes[rackNumber] = [
            {
              guid: positionConfig.guid,
              value: formattedPositions
            }
          ];

          Logger.log('  ' + rackNumber + ' positions: ' + formattedPositions);
        }

        bomOptions.bomAttributes = bomAttributes;
      } else {
        Logger.log('✓ Position tracking DISABLED - skipping BOM position attributes');
      }

      syncBOMToArena(client, rowItemGuid, bomLines, bomOptions);
      Logger.log('Added ' + bomLines.length + ' racks to row BOM');

      rowItems.push({
        rowNumber: row.rowNumber,
        sheetRow: row.sheetRow,
        itemNumber: rowItemNumber,
        guid: rowItemGuid,
        name: rowName,
        positions: positionNames
      });

    } catch (error) {
      Logger.log('Error creating row item: ' + error.message);
      ui.alert('Error', 'Failed to create row item: ' + error.message, ui.ButtonSet.OK);
      return null;
    }
  }

  return rowItems;
}

/**
 * Creates POD item in Arena with all rows as BOM
 * @param {Array} rowItems - Array of row item objects
 * @param {string} podCategory - Category to use for POD item
 * @return {Object} Created POD item metadata
 */
function createPODItem(rowItems, podCategory) {
  var ui = SpreadsheetApp.getUi();
  var client = new ArenaAPIClient();

  // Prompt user for POD name
  var promptMsg = '========================================\n' +
                  'CREATING POD ITEM (Top-Level Assembly)\n' +
                  '========================================\n\n' +
                  'This POD will contain ' + rowItems.length + ' Row item(s) in its BOM:\n\n';

  rowItems.forEach(function(row, index) {
    promptMsg += '  ' + (index + 1) + '. ' + row.name + ' (' + row.itemNumber + ')\n';
  });

  promptMsg += '\n----------------------------------------\n';
  promptMsg += 'Enter a name for this POD item in Arena:\n';
  promptMsg += '(e.g., "Data Center Pod A", "West Wing POD")';

  var nameResponse = ui.prompt('Create POD Item', promptMsg, ui.ButtonSet.OK_CANCEL);

  if (nameResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Error', 'POD creation cancelled.', ui.ButtonSet.OK);
    return null;
  }

  var podName = nameResponse.getResponseText().trim();
  if (!podName) {
    ui.alert('Error', 'POD name is required.', ui.ButtonSet.OK);
    return null;
  }

  try {
    Logger.log('=== CREATING POD ITEM ===');
    Logger.log('POD name: ' + podName);
    Logger.log('POD category (should be GUID): ' + podCategory);
    Logger.log('POD category type: ' + typeof podCategory);
    Logger.log('POD description: Point of Delivery containing ' + rowItems.length + ' rows');

    // Arena API expects category as an object with guid, not a simple string
    var createPODPayload = {
      name: podName,
      category: {
        guid: podCategory
      },
      description: 'Point of Delivery containing ' + rowItems.length + ' rows'
    };

    Logger.log('Full createItem payload for POD: ' + JSON.stringify(createPODPayload));

    // Create POD item in Arena
    var podItem = client.createItem(createPODPayload);

    var podItemGuid = podItem.guid || podItem.Guid;
    var podItemNumber = podItem.number || podItem.Number;

    Logger.log('Created POD item: ' + podItemNumber + ' (GUID: ' + podItemGuid + ')');

    // Handle manual item numbering if needed
    if (!podItemNumber) {
      Logger.log('⚠ Item number is null - category may not have auto-numbering enabled');

      var numberPrompt = '========================================\n' +
                         'ITEM NUMBER REQUIRED\n' +
                         '========================================\n\n' +
                         'The POD category does not have auto-numbering enabled.\n' +
                         'Please enter an item number for this POD:\n\n' +
                         'POD: ' + podName + '\n' +
                         'GUID: ' + podItemGuid;

      var numberResponse = ui.prompt('Enter Item Number', numberPrompt, ui.ButtonSet.OK_CANCEL);

      if (numberResponse.getSelectedButton() !== ui.Button.OK) {
        throw new Error('Item number is required but was not provided');
      }

      podItemNumber = numberResponse.getResponseText().trim();
      if (!podItemNumber) {
        throw new Error('Item number cannot be empty');
      }

      // Update the item with the user-provided number
      Logger.log('Updating POD item with manual number: ' + podItemNumber);
      client.updateItem(podItemGuid, { number: podItemNumber });
      Logger.log('✓ POD item number set to: ' + podItemNumber);
    }

    // Create BOM for POD (add all rows with quantity 1)
    var bomLines = rowItems.map(function(row) {
      return {
        itemNumber: row.itemNumber,
        itemGuid: row.guid,  // Use GUID from created row items
        quantity: 1,
        level: 0
      };
    });

    syncBOMToArena(client, podItemGuid, bomLines);
    Logger.log('Added ' + bomLines.length + ' rows to POD BOM');

    return {
      itemNumber: podItemNumber,
      guid: podItemGuid,
      name: podName
    };

  } catch (error) {
    Logger.log('Error creating POD item: ' + error.message);
    ui.alert('Error', 'Failed to create POD item: ' + error.message, ui.ButtonSet.OK);
    return null;
  }
}

/**
 * Attempts to rollback (delete) created items in case of failure
 * Deletes in reverse order: POD → Rows → Racks
 * @param {Object} context - Creation context with createdItems tracking
 * @param {ArenaAPIClient} client - Arena API client
 * @return {Object} {success: boolean, deletedCount: number, errors: Array}
 */
function attemptRollback(context, client) {
  Logger.log('========================================');
  Logger.log('ROLLBACK - ATTEMPTING CLEANUP');
  Logger.log('========================================');

  if (!context || !context.createdItems) {
    Logger.log('No context or createdItems to rollback');
    return { success: true, deletedCount: 0, errors: [] };
  }

  var deletedCount = 0;
  var errors = [];
  var createdItems = context.createdItems;

  // Delete in reverse order: POD → Rows → Racks
  var deleteOrder = ['POD', 'Row', 'Rack'];

  for (var typeIdx = 0; typeIdx < deleteOrder.length; typeIdx++) {
    var itemType = deleteOrder[typeIdx];
    var itemsOfType = [];

    // Collect all items of this type
    for (var i = 0; i < createdItems.length; i++) {
      if (createdItems[i].type === itemType) {
        itemsOfType.push(createdItems[i]);
      }
    }

    if (itemsOfType.length === 0) {
      Logger.log('No ' + itemType + ' items to delete');
      continue;
    }

    Logger.log('Deleting ' + itemsOfType.length + ' ' + itemType + ' item(s)...');

    for (var j = 0; j < itemsOfType.length; j++) {
      var item = itemsOfType[j];
      try {
        Logger.log('  Deleting ' + itemType + ': ' + item.itemNumber + ' (GUID: ' + item.guid + ')');

        // Arena API: DELETE /items/{guid}
        client.makeRequest('/items/' + item.guid, { method: 'DELETE' });

        deletedCount++;
        Logger.log('  ✓ Deleted: ' + item.itemNumber);

        // Small delay to avoid rate limiting
        Utilities.sleep(200);
      } catch (deleteError) {
        var errMsg = 'Failed to delete ' + itemType + ' ' + item.itemNumber + ': ' + deleteError.message;
        Logger.log('  ❌ ' + errMsg);
        errors.push(errMsg);
        // Continue trying to delete other items
      }
    }
  }

  Logger.log('========================================');
  Logger.log('ROLLBACK - COMPLETE');
  Logger.log('Deleted: ' + deletedCount + ' items');
  Logger.log('Errors: ' + errors.length);
  Logger.log('========================================');

  return {
    success: errors.length === 0,
    deletedCount: deletedCount,
    errors: errors
  };
}

/**
 * Validates all preconditions before starting POD push
 * Checks Arena connection, sheet structure, components exist, attributes configured
 * @param {Sheet} overviewSheet - Overview sheet to validate
 * @param {Array} customRacks - Array of custom rack objects to validate
 * @return {Object} {success: boolean, errors: Array, warnings: Array}
 */
function validatePreconditions(overviewSheet, customRacks) {
  Logger.log('========================================');
  Logger.log('PRE-FLIGHT VALIDATION - START');
  Logger.log('========================================');

  var errors = [];
  var warnings = [];
  var client;

  // 1. Validate Arena connection
  Logger.log('1. Validating Arena connection...');
  try {
    client = new ArenaAPIClient();
    var testEndpoint = client.makeRequest('/settings/items/attributes', { method: 'GET' });
    if (!testEndpoint) {
      errors.push('Arena connection test failed - no response from API');
    } else {
      Logger.log('✓ Arena connection successful');
    }
  } catch (connError) {
    errors.push('Arena connection failed: ' + connError.message);
    // Can't continue without connection
    return {
      success: false,
      errors: errors,
      warnings: warnings
    };
  }

  // 2. Validate overview sheet structure
  Logger.log('2. Validating overview sheet structure...');
  try {
    if (!overviewSheet) {
      errors.push('Overview sheet is null or undefined');
    } else {
      var data = overviewSheet.getDataRange().getValues();
      if (data.length === 0) {
        errors.push('Overview sheet is empty');
      } else {
        // Check for position headers
        var hasPositionHeaders = false;
        for (var i = 0; i < data.length && i < 10; i++) {
          for (var j = 0; j < data[i].length; j++) {
            var cell = data[i][j];
            if (cell && cell.toString().toLowerCase().indexOf('pos') === 0) {
              hasPositionHeaders = true;
              break;
            }
          }
          if (hasPositionHeaders) break;
        }

        if (!hasPositionHeaders) {
          errors.push('Overview sheet missing position headers (e.g., "Pos 1", "Pos 2", ...)');
        } else {
          Logger.log('✓ Overview sheet structure valid');
        }
      }
    }
  } catch (sheetError) {
    errors.push('Error validating overview sheet: ' + sheetError.message);
  }

  // 3. Validate BOM position attribute (if configured)
  try {
    var positionConfig = getBOMPositionAttributeConfig();
    if (positionConfig) {
      Logger.log('3. Validating BOM Position attribute...');
      Logger.log('   BOM Position attribute configured: ' + positionConfig.name);

      // Try to fetch the attribute to ensure it exists in Arena
      try {
        var bomAttrs = getBOMAttributes(client);
        var attrFound = false;
        for (var i = 0; i < bomAttrs.length; i++) {
          if (bomAttrs[i].guid === positionConfig.guid) {
            attrFound = true;
            Logger.log('   ✓ Found BOM attribute: ' + bomAttrs[i].name + ' (GUID: ' + bomAttrs[i].guid + ')');
            break;
          }
        }

        if (!attrFound) {
          Logger.log('   ⚠ BOM Position attribute not found - clearing invalid configuration');
          // Clear the invalid configuration so it won't be used during row creation
          clearBOMPositionAttribute();

          warnings.push('BOM Position attribute "' + positionConfig.name + '" (GUID: ' + positionConfig.guid + ') configured but not found in Arena BOM attributes. ' +
                       'It may have been deleted or is an Item attribute (not a BOM attribute). ' +
                       'Configuration has been cleared automatically. ' +
                       'Reconfigure using: Arena Data Center > Configuration > Rack BOM Location Setting');
        } else {
          Logger.log('   ✓ BOM Position attribute "' + positionConfig.name + '" exists in Arena');
        }
      } catch (bomAttrError) {
        Logger.log('   ⚠ Could not fetch BOM attributes - clearing invalid configuration');
        // If we can't fetch BOM attributes, clear the config to prevent errors during row creation
        clearBOMPositionAttribute();

        warnings.push('Could not verify BOM Position attribute: ' + bomAttrError.message + '. ' +
                     'Configuration has been cleared automatically. Position tracking will be disabled for this POD push.');
      }
    } else {
      Logger.log('3. BOM Position attribute - DISABLED (position tracking will be skipped)');
    }
  } catch (bomConfigError) {
    warnings.push('Error checking BOM Position config: ' + bomConfigError.message);
  }

  // 5. Validate all child components for custom racks exist in Arena
  Logger.log('5. Validating child components for custom racks...');
  if (customRacks && customRacks.length > 0) {
    var allChildNumbers = [];
    var childLookupMap = {}; // Track which racks need which children

    for (var r = 0; r < customRacks.length; r++) {
      var rack = customRacks[r];
      try {
        var children = getRackConfigChildren(rack.sheet);

        if (!children || children.length === 0) {
          warnings.push('Rack "' + rack.itemNumber + '" has no child components (empty BOM)');
        } else {
          for (var c = 0; c < children.length; c++) {
            var childNumber = children[c].itemNumber;
            if (allChildNumbers.indexOf(childNumber) === -1) {
              allChildNumbers.push(childNumber);
            }

            // Track which rack needs this child
            if (!childLookupMap[childNumber]) {
              childLookupMap[childNumber] = [];
            }
            if (childLookupMap[childNumber].indexOf(rack.itemNumber) === -1) {
              childLookupMap[childNumber].push(rack.itemNumber);
            }
          }
        }
      } catch (childError) {
        errors.push('Error reading children for rack "' + rack.itemNumber + '": ' + childError.message);
      }
    }

    // Now validate each child component exists in Arena
    Logger.log('Checking ' + allChildNumbers.length + ' unique child components in Arena...');
    var missingComponents = [];

    for (var i = 0; i < allChildNumbers.length; i++) {
      var childNum = allChildNumbers[i];
      try {
        Logger.log('  Checking component: ' + childNum);
        var childItem = client.getItemByNumber(childNum);

        if (!childItem) {
          var racksNeedingThis = childLookupMap[childNum].join(', ');
          missingComponents.push(childNum + ' (needed by: ' + racksNeedingThis + ')');
        } else {
          Logger.log('  ✓ Found: ' + childNum);
        }

        // Small delay to avoid rate limiting
        Utilities.sleep(50);
      } catch (lookupError) {
        var racksNeedingThis = childLookupMap[childNum].join(', ');
        missingComponents.push(childNum + ' (needed by: ' + racksNeedingThis + ') - Error: ' + lookupError.message);
      }
    }

    if (missingComponents.length > 0) {
      errors.push('Missing child components in Arena (' + missingComponents.length + ' total):\n  • ' +
                  missingComponents.join('\n  • '));
    } else {
      Logger.log('✓ All ' + allChildNumbers.length + ' child components found in Arena');
    }
  } else {
    Logger.log('✓ No custom racks to validate');
  }

  // 6. Summary
  Logger.log('========================================');
  Logger.log('PRE-FLIGHT VALIDATION - COMPLETE');
  Logger.log('========================================');
  Logger.log('Errors: ' + errors.length);
  Logger.log('Warnings: ' + warnings.length);

  if (errors.length > 0) {
    Logger.log('ERRORS:');
    errors.forEach(function(err) {
      Logger.log('  ❌ ' + err);
    });
  }

  if (warnings.length > 0) {
    Logger.log('WARNINGS:');
    warnings.forEach(function(warn) {
      Logger.log('  ⚠ ' + warn);
    });
  }

  return {
    success: errors.length === 0,
    errors: errors,
    warnings: warnings
  };
}

/**
 * NEW: Wizard-based POD push - scans sheets once and presents comprehensive UI
 * This is the new entry point that replaces the old dialog-by-dialog approach
 */
function pushPODStructureToArenaNew() {
  Logger.log('==========================================');
  Logger.log('POD PUSH WIZARD - START');
  Logger.log('==========================================');

  // Show loading modal immediately - it will call back to prepare data
  var html = HtmlService.createHtmlOutputFromFile('PODPushLoadingModal')
    .setWidth(400)
    .setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(html, 'Preparing POD Structure');
}

/**
 * Prepares all data for the POD push wizard by scanning sheets ONCE
 * Called from the loading modal, returns comprehensive data structure for the UI
 */
function preparePODWizardDataForModal() {
  Logger.log('Preparing POD wizard data...');

  var client = new ArenaAPIClient();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Find overview sheet
  var overviewSheet = null;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase().indexOf('overview') !== -1) {
      overviewSheet = sheets[i];
      break;
    }
  }

  if (!overviewSheet) {
    return { success: false, message: 'Overview sheet not found.' };
  }

  // Scan overview for rack/row structure
  var overviewData = scanOverviewByRow(overviewSheet);

  if (overviewData.length === 0) {
    return { success: false, message: 'No racks found in overview sheet.' };
  }

  // Get all unique rack numbers
  var allRackNumbers = [];
  overviewData.forEach(function(row) {
    row.positions.forEach(function(pos) {
      if (allRackNumbers.indexOf(pos.itemNumber) === -1) {
        allRackNumbers.push(pos.itemNumber);
      }
    });
  });

  Logger.log('Found ' + allRackNumbers.length + ' unique racks in overview: ' + allRackNumbers.join(', '));

  // Build rack config map (scan sheets ONCE)
  var rackConfigMap = {};
  sheets.forEach(function(sheet) {
    var metadata = getRackConfigMetadata(sheet);
    if (metadata) {
      Logger.log('Found rack config sheet: ' + metadata.itemNumber + ' - ' + metadata.itemName);
      rackConfigMap[metadata.itemNumber] = {
        sheet: sheet,
        metadata: metadata,
        childCount: Math.max(0, sheet.getLastRow() - 2)
      };
    }
  });

  var rackConfigNumbers = Object.keys(rackConfigMap);
  Logger.log('Built rack config map with ' + rackConfigNumbers.length + ' racks: ' + rackConfigNumbers.join(', '));

  // Separate placeholder vs existing racks
  var placeholderRacks = [];
  var existingRacks = [];

  allRackNumbers.forEach(function(itemNumber) {
    var rackConfig = rackConfigMap[itemNumber];

    if (!rackConfig) {
      Logger.log('⚠ No rack config found for: ' + itemNumber);
      return;
    }

    // Check if exists in Arena
    var existsInArena = false;
    var arenaItem = null;

    Logger.log('Checking if rack ' + itemNumber + ' exists in Arena...');

    try {
      arenaItem = client.getItemByNumber(itemNumber);

      Logger.log('Arena API response for ' + itemNumber + ': ' + JSON.stringify(arenaItem));

      if (arenaItem && (arenaItem.guid || arenaItem.Guid)) {
        // Exists in Arena (has valid GUID)
        existsInArena = true;
        var guid = arenaItem.guid || arenaItem.Guid;
        var name = arenaItem.name || arenaItem.Name || rackConfig.metadata.itemName;
        Logger.log('✓ Rack ' + itemNumber + ' EXISTS in Arena (GUID: ' + guid + ', Name: ' + name + ')');
      } else {
        Logger.log('⚠ Arena returned response but no GUID found for ' + itemNumber);
      }
    } catch (error) {
      // Error fetching = doesn't exist in Arena
      Logger.log('✗ Rack ' + itemNumber + ' NOT FOUND in Arena: ' + error.message);
    }

    if (existsInArena) {
      // Extract category info from Arena item
      var categoryName = '';
      if (arenaItem.category || arenaItem.Category) {
        var cat = arenaItem.category || arenaItem.Category;
        categoryName = cat.name || cat.Name || '';
      }

      existingRacks.push({
        itemNumber: itemNumber,
        name: arenaItem.name || arenaItem.Name || rackConfig.metadata.itemName,
        description: arenaItem.description || arenaItem.Description || rackConfig.metadata.description || '',
        category: categoryName,
        childCount: rackConfig.childCount,
        guid: arenaItem.guid || arenaItem.Guid
      });
      Logger.log('→ Added ' + itemNumber + ' to EXISTING racks list (category: ' + categoryName + ')');
    } else {
      // Doesn't exist in Arena - placeholder
      Logger.log('→ Adding ' + itemNumber + ' to PLACEHOLDER list');
      placeholderRacks.push({
        itemNumber: itemNumber,
        name: rackConfig.metadata.itemName || '',
        description: rackConfig.metadata.description || '',
        category: null,  // User will select
        childCount: rackConfig.childCount,
        sheet: rackConfig.sheet
      });
    }
  });

  Logger.log('Placeholder racks: ' + placeholderRacks.length);
  Logger.log('Existing racks: ' + existingRacks.length);

  // Prepare row data
  var rowsData = overviewData.map(function(row) {
    return {
      rowNumber: row.rowNumber,
      name: 'ROW-' + row.rowNumber,
      category: null,  // User will select
      positions: row.positions.map(function(pos) {
        return {
          position: pos.position,
          itemNumber: pos.itemNumber
        };
      })
    };
  });

  return {
    success: true,
    racks: placeholderRacks,
    existingRacks: existingRacks,
    rows: rowsData,
    pod: {
      name: '',
      category: null
    }
  };
}

/**
 * Shows the POD Push Wizard HTML dialog
 * This is the legacy version - kept for backwards compatibility
 */
function showPODPushWizard(wizardData) {
  var html = HtmlService.createHtmlOutputFromFile('PODPushWizard')
    .setWidth(1200)
    .setHeight(800);

  // Pass data to wizard
  var scriptlet = '<script>initializeWizard(' + JSON.stringify(wizardData) + ');</script>';
  var content = html.getContent() + scriptlet;
  html.setContent(content);

  SpreadsheetApp.getUi().showModalDialog(html, 'POD Structure Push Wizard');
}

/**
 * Shows the POD Push Wizard with pre-prepared data
 * Called from the loading modal after data is ready
 */
function showPODPushWizardWithData(wizardData) {
  Logger.log('Showing wizard with prepared data...');

  if (!wizardData.success) {
    SpreadsheetApp.getUi().alert('Error', wizardData.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var html = HtmlService.createHtmlOutputFromFile('PODPushWizard')
    .setWidth(1200)
    .setHeight(800);

  // Pass data to wizard
  var scriptlet = '<script>initializeWizard(' + JSON.stringify(wizardData) + ');</script>';
  var content = html.getContent() + scriptlet;
  html.setContent(content);

  SpreadsheetApp.getUi().showModalDialog(html, 'POD Structure Push Wizard');
}

/**
 * Transitions from loading modal to wizard
 * Called by loading modal when data is ready
 * Opens wizard then signals back so loading modal can close
 */
function transitionToWizard(wizardData) {
  Logger.log('Transitioning from loading modal to wizard...');

  // Open the wizard with prepared data
  showPODPushWizardWithData(wizardData);

  // Return success so loading modal knows wizard is open and can close itself
  return { success: true };
}

/**
 * Executes the batch POD push with data from the wizard
 * This is called by the wizard after user fills in all data
 */
function executePODPush(wizardData) {
  Logger.log('==========================================');
  Logger.log('EXECUTING BATCH POD PUSH');
  Logger.log('==========================================');

  var client = new ArenaAPIClient();
  var createdRacks = [];
  var createdRows = [];

  try {
    // STEP 1: Create all placeholder racks (batch)
    Logger.log('Step 1: Creating ' + wizardData.racks.length + ' placeholder racks...');

    for (var i = 0; i < wizardData.racks.length; i++) {
      var rack = wizardData.racks[i];

      Logger.log('Creating rack: ' + rack.itemNumber);

      // Create item in Arena (without number initially)
      var newItem = client.createItem({
        name: rack.name,
        category: {
          guid: rack.category.guid
        },
        description: rack.description
      });

      var newItemGuid = newItem.guid || newItem.Guid;

      // Update with desired rack number
      client.updateItem(newItemGuid, { number: rack.itemNumber });

      // Get rack children and sync BOM
      var children = getRackConfigChildren(rack.sheet);
      var bomLines = [];

      for (var j = 0; j < children.length; j++) {
        var child = children[j];
        var childItem = client.getItemByNumber(child.itemNumber);
        var childGuid = childItem.guid || childItem.Guid;

        bomLines.push({
          itemNumber: child.itemNumber,
          itemGuid: childGuid,
          quantity: child.quantity || 1,
          level: 0
        });
      }

      syncBOMToArena(client, newItemGuid, bomLines);

      // Update rack status
      updateRackSheetStatus(rack.sheet, RACK_STATUS.SYNCED, newItemGuid, {
        changesSummary: 'Rack created in Arena via POD push',
        details: 'Created with ' + bomLines.length + ' BOM items'
      });

      createdRacks.push({
        itemNumber: rack.itemNumber,
        guid: newItemGuid,
        name: rack.name
      });

      Logger.log('✓ Created rack: ' + rack.itemNumber);
    }

    // STEP 2: Create all rows (batch)
    Logger.log('Step 2: Creating ' + wizardData.rows.length + ' rows...');

    for (var r = 0; r < wizardData.rows.length; r++) {
      var row = wizardData.rows[r];

      Logger.log('Creating row: ' + row.name);

      // Create row item
      var rowItem = client.createItem({
        name: row.name,
        category: {
          guid: row.category.guid
        },
        description: 'Row ' + row.rowNumber + ' containing ' + row.positions.length + ' racks'
      });

      var rowItemGuid = rowItem.guid || rowItem.Guid;
      var rowItemNumber = rowItem.number || rowItem.Number;

      // Create BOM for row (all racks at positions)
      var rowBomLines = [];
      for (var p = 0; p < row.positions.length; p++) {
        var pos = row.positions[p];
        var rackItem = client.getItemByNumber(pos.itemNumber);
        var rackGuid = rackItem.guid || rackItem.Guid;

        rowBomLines.push({
          itemNumber: pos.itemNumber,
          itemGuid: rackGuid,
          quantity: 1,
          level: 0,
          position: pos.position
        });
      }

      syncBOMToArena(client, rowItemGuid, rowBomLines);

      createdRows.push({
        itemNumber: rowItemNumber,
        guid: rowItemGuid,
        name: row.name
      });

      Logger.log('✓ Created row: ' + rowItemNumber);
    }

    // STEP 3: Create POD
    Logger.log('Step 3: Creating POD...');

    var podItem = client.createItem({
      name: wizardData.pod.name,
      category: {
        guid: wizardData.pod.category.guid
      },
      description: 'Point of Delivery containing ' + createdRows.length + ' rows'
    });

    var podItemGuid = podItem.guid || podItem.Guid;
    var podItemNumber = podItem.number || podItem.Number;

    // Create BOM for POD (all rows)
    var podBomLines = createdRows.map(function(row) {
      return {
        itemNumber: row.itemNumber,
        itemGuid: row.guid,
        quantity: 1,
        level: 0
      };
    });

    syncBOMToArena(client, podItemGuid, podBomLines);

    Logger.log('✓ Created POD: ' + podItemNumber);

    return {
      success: true,
      racksCreated: createdRacks.length,
      rowsCreated: createdRows.length,
      podItemNumber: podItemNumber,
      podGuid: podItemGuid
    };

  } catch (error) {
    Logger.log('ERROR in batch POD push: ' + error.message);
    throw error;
  }
}

/**
 * Shows category selector from wizard context
 * Called by PODPushWizard.html
 */
function showCategorySelector() {
  var html = HtmlService.createHtmlOutputFromFile('CategorySelector')
    .setWidth(600)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, 'Select Category');

  // Note: Category selection is handled via returnCategorySelection()
  // which will call back to the wizard
}

/**
 * OLD: Original dialog-by-dialog POD push (DEPRECATED - use pushPODStructureToArenaNew instead)
 * Main function to push POD structure to Arena
 * Creates POD -> Rows -> Racks hierarchy in Arena
 */
function pushPODStructureToArena() {
  Logger.log('==========================================');
  Logger.log('PUSH POD STRUCTURE TO ARENA - START');
  Logger.log('==========================================');

  var ui = SpreadsheetApp.getUi();
  var client = null;

  // Context object to track created items for potential rollback
  var context = {
    createdItems: []  // Array of {type: 'Rack'|'Row'|'POD', itemNumber: string, guid: string}
  };

  try {
    // Initialize Arena client for rollback if needed
    client = new ArenaAPIClient();
    Logger.log('Step 0: Finding overview sheet...');
    // Find overview sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var overviewSheet = null;

    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().toLowerCase().indexOf('overview') !== -1) {
        overviewSheet = sheets[i];
        break;
      }
    }

    if (!overviewSheet) {
      ui.alert('Error', 'Overview sheet not found.', ui.ButtonSet.OK);
      return;
    }

    // Step 1: Scan overview for racks
    var overviewData = scanOverviewByRow(overviewSheet);

    if (overviewData.length === 0) {
      ui.alert('Error', 'No racks found in overview sheet.', ui.ButtonSet.OK);
      return;
    }

    // Get unique rack item numbers
    var allRackNumbers = [];
    overviewData.forEach(function(row) {
      row.positions.forEach(function(pos) {
        if (allRackNumbers.indexOf(pos.itemNumber) === -1) {
          allRackNumbers.push(pos.itemNumber);
        }
      });
    });

    Logger.log('Found ' + allRackNumbers.length + ' unique racks in overview');

    // Step 2: Identify custom racks
    Logger.log('Step 2: Identifying custom racks...');
    ui.alert('Checking Racks', 'Checking which racks need to be created in Arena...', ui.ButtonSet.OK);
    var customRacks = identifyCustomRacks(allRackNumbers);

    // Step 3: Create custom rack items FIRST (before validation)
    // This ensures all racks exist before we try to create rows that reference them
    if (customRacks.length > 0) {
      Logger.log('Step 3: Creating ' + customRacks.length + ' custom rack item(s)...');
      var msg = 'Found ' + customRacks.length + ' custom rack(s) that need Arena items:\n\n';
      customRacks.forEach(function(rack) {
        msg += '  - ' + rack.itemNumber + '\n';
      });
      msg += '\nThese will be created automatically before building the POD structure.';

      ui.alert('Custom Racks Detected', msg, ui.ButtonSet.OK);

      var rackResult = createCustomRackItems(customRacks);

      if (!rackResult.success) {
        ui.alert('Error', 'Failed to create custom racks: ' + rackResult.message, ui.ButtonSet.OK);
        return;
      }

      // Track created racks for rollback support
      if (rackResult.createdItems) {
        for (var ri = 0; ri < rackResult.createdItems.length; ri++) {
          context.createdItems.push({
            type: 'Rack',
            itemNumber: rackResult.createdItems[ri].itemNumber,
            guid: rackResult.createdItems[ri].guid
          });
        }
        Logger.log('Tracked ' + rackResult.createdItems.length + ' rack(s) in context for rollback');
      }

      ui.alert('Success', 'Created ' + rackResult.createdItems.length + ' rack item(s) in Arena.\n\nNow running validation...', ui.ButtonSet.OK);

      // Step 3.5: Verify newly created racks are findable in Arena
      Logger.log('Step 3.5: Verifying newly created racks are findable...');
      var verificationErrors = [];

      for (var vi = 0; vi < rackResult.createdItems.length; vi++) {
        var createdRack = rackResult.createdItems[vi];
        var maxRetries = 3;
        var retryDelay = 1000; // 1 second
        var found = false;

        for (var attempt = 1; attempt <= maxRetries; attempt++) {
          try {
            Logger.log('  Verifying rack ' + createdRack.itemNumber + ' (attempt ' + attempt + '/' + maxRetries + ')');
            var verifyItem = client.getItemByNumber(createdRack.itemNumber);

            if (verifyItem && verifyItem.guid) {
              Logger.log('  ✓ Rack ' + createdRack.itemNumber + ' verified (GUID: ' + verifyItem.guid + ')');
              found = true;
              break;
            }
          } catch (verifyError) {
            Logger.log('  Attempt ' + attempt + ' failed: ' + verifyError.message);
            if (attempt < maxRetries) {
              Logger.log('  Waiting ' + retryDelay + 'ms before retry...');
              Utilities.sleep(retryDelay);
              retryDelay *= 2; // Exponential backoff
            }
          }
        }

        if (!found) {
          verificationErrors.push('Rack ' + createdRack.itemNumber + ' was created but could not be verified in Arena');
        }
      }

      if (verificationErrors.length > 0) {
        var verifyMsg = 'WARNING: Some newly created racks could not be verified:\n\n';
        verificationErrors.forEach(function(err) {
          verifyMsg += '• ' + err + '\n';
        });
        verifyMsg += '\nThis may cause issues when creating rows. Continue anyway?';

        var verifyResponse = ui.alert('Verification Warning', verifyMsg, ui.ButtonSet.YES_NO);
        if (verifyResponse !== ui.Button.YES) {
          ui.alert('Cancelled', 'POD creation cancelled.', ui.ButtonSet.OK);
          return;
        }
      } else {
        Logger.log('✓ All newly created racks verified successfully');
      }
    } else {
      Logger.log('Step 3: No custom racks need to be created');
    }

    // Step 4: PRE-FLIGHT VALIDATION - Now that all racks exist, validate everything
    Logger.log('Step 4: Running pre-flight validation...');
    ui.alert('Validation', 'Running comprehensive pre-flight validation...\n\nThis will check:\n• Arena connection\n• Overview sheet structure\n• BOM Position attribute (if configured)\n• All child components exist', ui.ButtonSet.OK);

    var validation = validatePreconditions(overviewSheet, customRacks);

    // Show warnings if any (non-blocking)
    if (validation.warnings.length > 0) {
      var warningMsg = 'Pre-flight validation found ' + validation.warnings.length + ' warning(s):\n\n';
      validation.warnings.forEach(function(warn, idx) {
        warningMsg += (idx + 1) + '. ' + warn + '\n';
      });
      warningMsg += '\nThese are warnings only. Continue anyway?';

      var warnResponse = ui.alert('Validation Warnings', warningMsg, ui.ButtonSet.YES_NO);
      if (warnResponse !== ui.Button.YES) {
        ui.alert('Cancelled', 'POD creation cancelled due to validation warnings.', ui.ButtonSet.OK);
        return;
      }
    }

    // Show errors and stop if any (blocking)
    if (!validation.success) {
      var errorMsg = 'Pre-flight validation FAILED with ' + validation.errors.length + ' error(s):\n\n';
      validation.errors.forEach(function(err, idx) {
        errorMsg += (idx + 1) + '. ' + err + '\n\n';
      });
      errorMsg += 'Please fix these issues and try again.\n\n';
      errorMsg += 'Check View → Executions for detailed logs.';

      ui.alert('Validation Failed', errorMsg, ui.ButtonSet.OK);
      Logger.log('Pre-flight validation failed. Aborting POD push.');
      return;
    }

    Logger.log('✓ Pre-flight validation passed!');
    ui.alert('Validation Passed', 'All pre-flight checks passed!\n\nReady to create POD structure.', ui.ButtonSet.OK);

    // Step 5: Show summary and confirm
    var summaryMsg = '========================================\n' +
                     'READY TO CREATE POD STRUCTURE\n' +
                     '========================================\n\n' +
                     'The following items will be created in Arena:\n\n' +
                     '1. ROW ITEMS (' + overviewData.length + ' total)\n';

    overviewData.forEach(function(row, index) {
      var rackCount = row.positions.length;
      summaryMsg += '   • Row ' + row.rowNumber + ' - Contains ' + rackCount + ' rack(s)\n';
    });

    summaryMsg += '\n2. POD ITEM (1 total)\n';
    summaryMsg += '   • Top-level assembly containing all ' + overviewData.length + ' rows\n\n';
    summaryMsg += '----------------------------------------\n';
    summaryMsg += 'You will be prompted to name each item.\n\n';
    summaryMsg += 'Continue?';

    var confirmResponse = ui.alert('Confirm POD Creation', summaryMsg, ui.ButtonSet.YES_NO);

    if (confirmResponse !== ui.Button.YES) {
      ui.alert('Cancelled', 'POD creation cancelled.', ui.ButtonSet.OK);
      return;
    }

    // Step 6: Prompt for Row item category
    var rowCategorySelection = showCategorySelector(
      'Category for Row Items',
      'Select the Arena category to use for all Row items (' + overviewData.length + ' rows)'
    );

    if (!rowCategorySelection) {
      ui.alert('Cancelled', 'POD creation cancelled.', ui.ButtonSet.OK);
      return;
    }

    var rowCategoryGuid = rowCategorySelection.guid;
    var rowCategoryName = rowCategorySelection.name;
    Logger.log('Selected Row category: ' + rowCategoryName + ' (GUID: ' + rowCategoryGuid + ')');

    // Step 7: Create Row items
    var rowItems = createRowItems(overviewData, rowCategoryGuid);

    if (!rowItems) {
      return; // Cancelled or error
    }

    // Track created rows for rollback support
    for (var rowIdx = 0; rowIdx < rowItems.length; rowIdx++) {
      context.createdItems.push({
        type: 'Row',
        itemNumber: rowItems[rowIdx].itemNumber,
        guid: rowItems[rowIdx].guid
      });
    }
    Logger.log('Tracked ' + rowItems.length + ' row(s) in context for rollback');

    // Step 8: Prompt for POD item category
    var podCategorySelection = showCategorySelector(
      'Category for POD Item',
      'Select the Arena category to use for the top-level POD assembly'
    );

    if (!podCategorySelection) {
      ui.alert('Cancelled', 'POD creation cancelled.', ui.ButtonSet.OK);
      return;
    }

    var podCategoryGuid = podCategorySelection.guid;
    var podCategoryName = podCategorySelection.name;
    Logger.log('Selected POD category: ' + podCategoryName + ' (GUID: ' + podCategoryGuid + ')');

    // Step 9: Create POD item
    ui.alert('Creating POD', 'Creating POD item in Arena...', ui.ButtonSet.OK);
    var podItem = createPODItem(rowItems, podCategoryGuid);

    if (!podItem) {
      return; // Cancelled or error
    }

    // Track created POD for rollback support (though at this point, success is almost certain)
    context.createdItems.push({
      type: 'POD',
      itemNumber: podItem.itemNumber,
      guid: podItem.guid
    });
    Logger.log('Tracked POD in context for rollback');

    // Step 10: Update overview sheet with POD/Row info
    updateOverviewWithPODInfo(overviewSheet, podItem, rowItems);

    // Fetch full POD item to get Arena URL
    var podArenaItem = client.getItemByNumber(podItem.itemNumber);
    var podArenaUrl = buildArenaItemURLFromItem(podArenaItem, podItem.itemNumber);

    // Build rack summary for success message
    var rackSummary = '';
    var rackCount = 0;
    rowItems.forEach(function(rowItem) {
      rowItem.racks.forEach(function(rack) {
        rackCount++;
        if (rackCount <= 5) {  // Show first 5 racks
          rackSummary += '  • ' + rack.itemNumber + ' - ' + rack.name + '\n';
        }
      });
    });

    if (rackCount > 5) {
      rackSummary += '  ... and ' + (rackCount - 5) + ' more rack(s)\n';
    }

    // Success message with clickable Arena link
    var successMsg = 'POD Structure Created Successfully!\n\n' +
                     '📦 POD Item: ' + podItem.name + '\n' +
                     '   Item #: ' + podItem.itemNumber + '\n' +
                     '   Category: ' + podCategoryName + '\n\n' +
                     '📊 Structure:\n' +
                     '   Rows: ' + rowItems.length + '\n' +
                     '   Racks: ' + rackCount + '\n\n';

    if (rackSummary) {
      successMsg += '🔧 Racks Included:\n' + rackSummary + '\n';
    }

    successMsg += '✓ Overview sheet updated with Arena links\n\n' +
                  '🔗 View in Arena:\n' + podArenaUrl;

    ui.alert('Success', successMsg, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('ERROR in pushPODStructureToArena: ' + error.message);
    Logger.log('Stack trace: ' + error.stack);

    // Attempt rollback if any items were created
    var rollbackMsg = '';
    if (context.createdItems.length > 0 && client) {
      Logger.log('Attempting rollback of ' + context.createdItems.length + ' created item(s)...');

      var rollbackChoice = ui.alert(
        'Error - Rollback?',
        'Failed to create POD structure: ' + error.message + '\n\n' +
        'Created items so far: ' + context.createdItems.length + '\n\n' +
        'Would you like to rollback (delete) the items that were created?\n\n' +
        'YES = Delete created items\n' +
        'NO = Keep items in Arena (you can continue manually)',
        ui.ButtonSet.YES_NO
      );

      if (rollbackChoice === ui.Button.YES) {
        ui.alert('Rolling Back', 'Deleting created items from Arena...', ui.ButtonSet.OK);

        var rollbackResult = attemptRollback(context, client);

        if (rollbackResult.success) {
          rollbackMsg = '\n\n✓ Rollback successful. Deleted ' + rollbackResult.deletedCount + ' item(s) from Arena.';
        } else {
          rollbackMsg = '\n\n⚠ Rollback partially failed.\n' +
                       'Deleted: ' + rollbackResult.deletedCount + ' item(s)\n' +
                       'Errors: ' + rollbackResult.errors.length + '\n\n' +
                       'Some items may still exist in Arena. Check View → Executions for details.';
        }
      } else {
        rollbackMsg = '\n\nRollback skipped. ' + context.createdItems.length + ' item(s) remain in Arena.';
      }
    }

    ui.alert('Error', 'Failed to create POD structure: ' + error.message + rollbackMsg, ui.ButtonSet.OK);
  }
}

/**
 * Updates overview sheet with POD and Row information
 * @param {Sheet} sheet - Overview sheet
 * @param {Object} podItem - POD item metadata
 * @param {Array} rowItems - Array of row item objects
 */
function updateOverviewWithPODInfo(sheet, podItem, rowItems) {
  // Insert column after row numbers for Row Item info
  var data = sheet.getDataRange().getValues();

  // Find header row
  var headerRow = -1;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    for (var j = 0; j < row.length; j++) {
      if (row[j] && row[j].toString().toLowerCase().indexOf('pos') === 0) {
        headerRow = i + 1; // Convert to 1-based
        break;
      }
    }
    if (headerRow !== -1) break;
  }

  if (headerRow === -1) return;

  // Insert new column at position 2 (after row numbers)
  sheet.insertColumnAfter(1);

  // Set header
  sheet.getRange(headerRow, 2).setValue('Row Item');
  sheet.getRange(headerRow, 2).setFontWeight('bold').setBackground('#f0f0f0');

  // Fetch full item data from Arena for building URLs
  var client = new ArenaAPIClient();

  // Add row item links
  rowItems.forEach(function(rowItem) {
    var rowArenaItem = client.getItemByNumber(rowItem.itemNumber);
    var arenaUrl = buildArenaItemURLFromItem(rowArenaItem, rowItem.itemNumber);
    var formula = '=HYPERLINK("' + arenaUrl + '", "' + rowItem.itemNumber + '")';
    sheet.getRange(rowItem.sheetRow, 2).setFormula(formula);
    sheet.getRange(rowItem.sheetRow, 2).setFontColor('#0000FF');
  });

  // Add POD info at top (in a merged cell above the grid)
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, 10).merge();
  var podArenaItem = client.getItemByNumber(podItem.itemNumber);
  var podUrl = buildArenaItemURLFromItem(podArenaItem, podItem.itemNumber);
  var podFormula = '=HYPERLINK("' + podUrl + '", "POD: ' + podItem.name + ' (' + podItem.itemNumber + ')")';
  sheet.getRange(1, 1).setFormula(podFormula);
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(12).setFontColor('#0000FF');

  Logger.log('Updated overview sheet with POD and Row information');
}

/**
 * Repairs existing POD/Row BOMs from overview sheet  
 * For when items were already created but BOMs are empty due to previous bugs
 */
function repairPODAndRowBOMs() {
  Logger.log('==========================================');
  Logger.log('REPAIR POD AND ROW BOMs - START');
  Logger.log('==========================================');

  var ui = SpreadsheetApp.getUi();
  var client = new ArenaAPIClient();

  try {
    // Find overview sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var overviewSheet = null;

    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().toLowerCase().indexOf('overview') !== -1) {
        overviewSheet = sheets[i];
        break;
      }
    }

    if (!overviewSheet) {
      ui.alert('Error', 'Overview sheet not found.', ui.ButtonSet.OK);
      return;
    }

    Logger.log('Found overview sheet: ' + overviewSheet.getName());

    // Step 1: Read POD and Row item numbers from sheet
    var data = overviewSheet.getDataRange().getValues();

    // Find POD item number (first row, should be hyperlink formula)
    var podItemNumber = null;
    var podRow = data[0][0];
    if (podRow && typeof podRow === 'string') {
      // Extract item number from format "POD: name (ITEM-NUMBER)"
      var match = podRow.match(/\(([^)]+)\)/);
      if (match) {
        podItemNumber = match[1];
      }
    }

    if (!podItemNumber) {
      ui.alert('Error', 'Could not find POD item number in overview sheet.\n\nPlease ensure the POD was created and the overview sheet has the POD link.', ui.ButtonSet.OK);
      return;
    }

    Logger.log('Found POD item number: ' + podItemNumber);

    // Step 2: Find header row and Row Item column
    var headerRow = -1;
    var rowItemCol = -1;
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      for (var j = 0; j < row.length; j++) {
        if (row[j] && row[j].toString().toLowerCase().indexOf('pos') === 0) {
          headerRow = i;
          // Row Item column should be right before Pos columns
          rowItemCol = j - 1;
          break;
        }
      }
      if (headerRow !== -1) break;
    }

    if (headerRow === -1 || rowItemCol === -1) {
      ui.alert('Error', 'Could not find row layout in overview sheet.', ui.ButtonSet.OK);
      return;
    }

    // Step 3: Extract Row item numbers and their rack placements
    var rowData = [];
    var headers = data[headerRow];
    var firstPosCol = rowItemCol + 1;

    for (var i = headerRow + 1; i < data.length; i++) {
      var row = data[i];
      var rowItemNumber = row[rowItemCol];

      // Skip if no row item number
      if (!rowItemNumber || typeof rowItemNumber !== 'string') continue;

      // Extract item number from hyperlink if needed
      if (rowItemNumber.indexOf('HYPERLINK') !== -1) {
        var match = rowItemNumber.match(/"([^"]+)"/g);
        if (match && match.length >= 2) {
          rowItemNumber = match[1].replace(/"/g, '');
        }
      }

      // Get rack placements for this row
      var rackNumbers = [];
      for (var j = firstPosCol; j < row.length; j++) {
        var cellValue = row[j];
        if (cellValue && typeof cellValue === 'string' && cellValue.trim() !== '') {
          rackNumbers.push(cellValue.trim());
        }
      }

      if (rackNumbers.length > 0) {
        rowData.push({
          rowItemNumber: rowItemNumber,
          racks: rackNumbers
        });
      }
    }

    Logger.log('Found ' + rowData.length + ' row items with rack placements');

    // Step 4: Show confirmation
    var confirmMsg = '========================================\n' +
                     'REPAIR BOM STRUCTURE\n' +
                     '========================================\n\n' +
                     'This will repair the following BOMs in Arena:\n\n' +
                     'POD Item: ' + podItemNumber + '\n' +
                     '  → Will add ' + rowData.length + ' Row items to POD BOM\n\n';

    rowData.forEach(function(row) {
      confirmMsg += 'Row Item: ' + row.rowItemNumber + '\n';
      confirmMsg += '  → Will add ' + row.racks.length + ' Rack items to Row BOM\n';
    });

    confirmMsg += '\n----------------------------------------\n';
    confirmMsg += 'This will DELETE any existing BOM lines and replace them.\n\n';
    confirmMsg += 'Continue?';

    var response = ui.alert('Confirm BOM Repair', confirmMsg, ui.ButtonSet.YES_NO);

    if (response !== ui.Button.YES) {
      ui.alert('Cancelled', 'BOM repair cancelled.', ui.ButtonSet.OK);
      return;
    }

    // Step 5: Look up POD GUID
    Logger.log('Looking up POD item: ' + podItemNumber);
    var podItem = client.getItemByNumber(podItemNumber);
    if (!podItem) {
      throw new Error('POD item not found in Arena: ' + podItemNumber);
    }
    var podGuid = podItem.guid || podItem.Guid;
    Logger.log('Found POD GUID: ' + podGuid);

    // Step 6: Repair Row BOMs
    var rowGuids = [];
    for (var i = 0; i < rowData.length; i++) {
      var row = rowData[i];

      Logger.log('Looking up Row item: ' + row.rowItemNumber);
      var rowItem = client.getItemByNumber(row.rowItemNumber);
      if (!rowItem) {
        Logger.log('WARNING: Row item not found: ' + row.rowItemNumber);
        continue;
      }
      var rowGuid = rowItem.guid || rowItem.Guid;
      Logger.log('Found Row GUID: ' + rowGuid);
      rowGuids.push({ number: row.rowItemNumber, guid: rowGuid });

      // Build BOM lines for this row (add racks)
      var bomLines = [];
      var rackCounts = {};

      // Count rack occurrences
      row.racks.forEach(function(rackNum) {
        if (!rackCounts[rackNum]) {
          rackCounts[rackNum] = 0;
        }
        rackCounts[rackNum]++;
      });

      // Look up GUIDs for each rack
      for (var rackNumber in rackCounts) {
        Logger.log('  Looking up Rack: ' + rackNumber);
        var rackItem = client.getItemByNumber(rackNumber);
        if (!rackItem) {
          Logger.log('  WARNING: Rack not found: ' + rackNumber);
          continue;
        }
        var rackGuid = rackItem.guid || rackItem.Guid;

        bomLines.push({
          itemNumber: rackNumber,
          itemGuid: rackGuid,
          quantity: rackCounts[rackNumber],
          level: 0
        });
        Logger.log('  Found Rack GUID: ' + rackNumber + ' → ' + rackGuid);
      }

      // Sync BOM to Arena
      Logger.log('Syncing BOM for Row: ' + row.rowItemNumber + ' (' + bomLines.length + ' racks)');
      syncBOMToArena(client, rowGuid, bomLines);
      Logger.log('✓ Repaired Row BOM: ' + row.rowItemNumber);
    }

    // Step 7: Repair POD BOM (add all rows)
    Logger.log('Repairing POD BOM: ' + podItemNumber);
    var podBomLines = rowGuids.map(function(row) {
      return {
        itemNumber: row.number,
        itemGuid: row.guid,
        quantity: 1,
        level: 0
      };
    });

    syncBOMToArena(client, podGuid, podBomLines);
    Logger.log('✓ Repaired POD BOM: ' + podItemNumber);

    // Success
    var successMsg = 'BOM Repair Complete!\n\n' +
                     'POD: ' + podItemNumber + '\n' +
                     '  → Added ' + rowGuids.length + ' rows to BOM\n\n';

    rowData.forEach(function(row) {
      var rackCount = row.racks.length;
      var uniqueRacks = {};
      row.racks.forEach(function(r) { uniqueRacks[r] = true; });
      successMsg += 'Row: ' + row.rowItemNumber + '\n';
      successMsg += '  → Added ' + Object.keys(uniqueRacks).length + ' unique rack types (' + rackCount + ' total)\n';
    });

    successMsg += '\n✓ All BOMs have been repaired in Arena!';

    ui.alert('Success', successMsg, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('Error repairing BOMs: ' + error.message + '\n' + error.stack);
    ui.alert('Error', 'Failed to repair BOMs:\n\n' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Builds a proper Arena web UI URL for an item
 * @param {Object} item - Arena item object (from API)
 * @param {string} itemNumber - Arena item number (fallback)
 * @return {string} Arena web UI URL
 */
function buildArenaItemURLFromItem(item, itemNumber) {
  try {
    if (!item) {
      Logger.log('WARNING: No item object provided for ' + itemNumber + ', using search URL');
      return 'https://app.bom.com/search?query=' + encodeURIComponent(itemNumber);
    }

    // Extract item_id from Arena item
    // Arena API may return these with different casing
    var itemId = item.itemId || item.ItemId || item.id || item.Id;

    // Log the full item structure to understand what Arena returns
    Logger.log('Item object keys for ' + itemNumber + ': ' + Object.keys(item).join(', '));

    // If we don't have the item ID, use search URL as fallback
    if (!itemId) {
      Logger.log('WARNING: Missing itemId for ' + itemNumber);
      Logger.log('Available properties: ' + JSON.stringify(Object.keys(item)));
      return 'https://app.bom.com/search?query=' + encodeURIComponent(itemNumber);
    }

    // Build the proper Arena web UI URL (item_id only, no version_id)
    var arenaUrl = 'https://app.bom.com/items/detail-spec?item_id=' + itemId;

    Logger.log('Built Arena URL for ' + itemNumber + ': ' + arenaUrl);
    return arenaUrl;

  } catch (error) {
    Logger.log('ERROR building Arena URL for ' + itemNumber + ': ' + error.message);
    // Fallback to search URL on error
    return 'https://app.bom.com/search?query=' + encodeURIComponent(itemNumber);
  }
}
