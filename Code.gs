/**
 * PTC Arena Sheets Data Center
 * Main entry point for Google Sheets Add-on
 */

/**
 * Runs when the add-on is installed
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Runs when the spreadsheet is opened
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Arena Data Center')
    .addSubMenu(ui.createMenu('Configuration')
      .addItem('Configure Arena Connection', 'showLoginWizard')
      .addItem('Configure Item Columns', 'showConfigureColumns')
      .addItem('Configure Category Colors', 'showConfigureColors')
      .addItem('Configure BOM Levels', 'showConfigureBOMLevels'))
    .addSeparator()
    .addItem('Show Item Picker', 'showItemPicker')
    .addSeparator()
    .addSubMenu(ui.createMenu('Create Layout')
      .addItem('New Tower Layout', 'createNewTowerLayout')
      .addItem('New Overview Layout', 'createNewOverviewLayout')
      .addItem('New Rack Configuration', 'createNewRackConfig')
      .addSeparator()
      .addItem('Auto-Link Racks to Overview', 'autoLinkRacksToOverviewAction'))
    .addSeparator()
    .addSubMenu(ui.createMenu('BOM Operations')
      .addItem('Pull BOM from Arena', 'pullBOMFromArena')
      .addItem('Push BOM to Arena', 'pushBOMToArena')
      .addSeparator()
      .addItem('Create Consolidated BOM', 'createConsolidatedBOMSheet'))
    .addSeparator()
    .addItem('Test Connection', 'testArenaConnection')
    .addItem('Clear Credentials', 'clearCredentials')
    .addToUi();
}

/**
 * Runs when a cell is edited
 * Used to trigger item insertion when user clicks a cell after selecting an item
 */
function onSelectionChange(e) {
  var selectedItem = getSelectedItem();

  if (selectedItem) {
    // User has an item selected from the picker and clicked a cell
    // Show a prompt to insert the item
    var sheet = SpreadsheetApp.getActiveSheet();
    var cell = sheet.getActiveCell();

    if (cell) {
      insertSelectedItem();
    }
  }
}

/**
 * Shows the login wizard for Arena API configuration
 */
function showLoginWizard() {
  var html = HtmlService.createHtmlOutputFromFile('LoginWizard')
    .setWidth(400)
    .setHeight(500)
    .setTitle('Configure Arena API Connection');
  SpreadsheetApp.getUi().showModalDialog(html, 'Arena API Configuration');
}

/**
 * Tests the Arena API connection
 */
function testArenaConnection() {
  var ui = SpreadsheetApp.getUi();

  if (!isAuthorized()) {
    ui.alert('Not Configured', 'Please configure your Arena API connection first.', ui.ButtonSet.OK);
    showLoginWizard();
    return;
  }

  try {
    var arenaClient = new ArenaAPIClient();
    var result = arenaClient.testConnection();

    if (result.success) {
      ui.alert('Success', 'Connection to Arena API successful!\n\nWorkspace: ' + getWorkspaceId(), ui.ButtonSet.OK);
    } else {
      ui.alert('Connection Failed', 'Could not connect to Arena API:\n' + result.error, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('Error', 'Connection test failed: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Imports data from Arena and populates all sheets
 */
function importArenaData() {
  var ui = SpreadsheetApp.getUi();

  if (!isAuthorized()) {
    ui.alert('Not Configured', 'Please configure your Arena API connection first.', ui.ButtonSet.OK);
    showLoginWizard();
    return;
  }

  // Confirm with user
  var response = ui.alert(
    'Import Arena Data',
    'This will fetch all items from Arena and populate the datacenter BOM sheets.\n\n' +
    'This may take a few minutes depending on the number of items.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    // Show progress message
    var progressMsg = 'Importing data from Arena...\n\n' +
                     'Step 1: Fetching items from Arena API...\n' +
                     'Please wait, this may take a few minutes.';

    showProgressDialog(progressMsg);

    // Create Arena API client
    var arenaClient = new ArenaAPIClient();

    // Fetch all items from Arena
    Logger.log('Starting Arena data import...');
    var arenaItems = arenaClient.getAllItems(100);

    if (!arenaItems || arenaItems.length === 0) {
      ui.alert('No Data', 'No items were found in Arena workspace.', ui.ButtonSet.OK);
      return;
    }

    Logger.log('Fetched ' + arenaItems.length + ' items from Arena');

    // Group items by rack type
    Logger.log('Step 2: Categorizing items by rack type...');
    var groupedItems = groupItemsByRackType(arenaItems);

    // Populate individual rack tabs
    Logger.log('Step 3: Populating rack tabs...');
    var rackResults = populateAllRackTabs(groupedItems);

    // Generate overhead layout
    Logger.log('Step 4: Generating overhead layout...');
    var overheadResult = generateOverheadLayout();

    // Generate Legend-NET summary
    Logger.log('Step 5: Generating Legend-NET summary...');
    var legendResult = updateNetworkingLegend();

    // Reorder sheets for better organization
    Logger.log('Step 6: Organizing sheets...');
    reorderSheets();

    // Build summary message
    var successCount = 0;
    var totalItems = 0;

    rackResults.forEach(function(result) {
      if (result.success) {
        successCount++;
        totalItems += result.rowsAdded || 0;
      }
    });

    var summaryMsg = 'Data import completed successfully!\n\n' +
                    'Items fetched from Arena: ' + arenaItems.length + '\n' +
                    'Rack tabs populated: ' + successCount + '\n' +
                    'Total items populated: ' + totalItems + '\n\n' +
                    'Sheets have been organized and formatted.';

    ui.alert('Import Complete', summaryMsg, ui.ButtonSet.OK);

    Logger.log('Arena data import completed successfully');

  } catch (error) {
    Logger.log('Error importing Arena data: ' + error.message);
    ui.alert('Import Error', 'An error occurred while importing data:\n\n' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Shows a progress dialog (HTML-based for better UX)
 * @param {string} message - Progress message to display
 */
function showProgressDialog(message) {
  var html = '<div style="padding: 20px; font-family: Arial, sans-serif;">' +
            '<h3 style="color: #1a73e8;">Arena Data Import</h3>' +
            '<p>' + message.replace(/\n/g, '<br>') + '</p>' +
            '<p style="color: #666; font-size: 12px;">This dialog will close automatically when complete.</p>' +
            '</div>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(200);

  // Note: This dialog will appear but won't update dynamically in Apps Script
  // For a better progress indicator, consider using a sidebar instead
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'Importing...');
}

/**
 * Clears stored credentials
 */
function clearCredentials() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Clear Credentials',
    'Are you sure you want to clear your Arena API credentials?',
    ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    clearArenaCredentials();
    ui.alert('Cleared', 'Arena API credentials have been cleared.', ui.ButtonSet.OK);
  }
}

/**
 * Shows the category color configuration dialog
 */
function showConfigureColors() {
  var html = HtmlService.createHtmlOutputFromFile('ConfigureColors')
    .setWidth(650)
    .setHeight(600)
    .setTitle('Configure Category Colors');
  SpreadsheetApp.getUi().showModalDialog(html, 'Category Colors');
}

/**
 * Loads category and color data for the configuration dialog
 * @return {Object} Object with categories and colors
 */
function loadCategoryColorData() {
  return {
    categories: getArenaCategories(),
    colors: getCategoryColors()
  };
}

/**
 * Shows the configure item columns dialog
 */
function showConfigureColumns() {
  var html = HtmlService.createHtmlOutputFromFile('ConfigureColumns')
    .setWidth(750)
    .setHeight(650)
    .setTitle('Configure Item Columns');
  SpreadsheetApp.getUi().showModalDialog(html, 'Item Columns');
}

/**
 * Loads column configuration data
 * @return {Object} Object with attributes, groups, and current selection
 */
function loadColumnConfigData() {
  var currentColumns = getItemColumns();
  var selectedMap = {};

  // Convert current columns to selection map
  currentColumns.forEach(function(col) {
    if (col.attributeGuid) {
      selectedMap[col.attributeGuid] = {
        guid: col.attributeGuid,
        name: col.attributeName || '',
        header: col.header || ''
      };
    }
  });

  return {
    attributes: getArenaAttributes(),
    groups: getAttributeGroups(),
    currentSelection: selectedMap
  };
}

/**
 * Gets saved attribute groups
 * @return {Array} Array of attribute groups
 */
function getAttributeGroups() {
  var json = PropertiesService.getScriptProperties().getProperty('ATTRIBUTE_GROUPS');
  if (!json) return [];
  try {
    return JSON.parse(json);
  } catch (e) {
    return [];
  }
}

/**
 * Saves an attribute group
 * @param {string} name - Group name
 * @param {Array} attributes - Array of attribute configurations
 * @return {Array} Updated list of groups
 */
function saveAttributeGroup(name, attributes) {
  var groups = getAttributeGroups();

  // Check if group exists
  var existingIndex = -1;
  for (var i = 0; i < groups.length; i++) {
    if (groups[i].name === name) {
      existingIndex = i;
      break;
    }
  }

  var newGroup = {
    name: name,
    attributes: attributes
  };

  if (existingIndex >= 0) {
    groups[existingIndex] = newGroup;
  } else {
    groups.push(newGroup);
  }

  PropertiesService.getScriptProperties().setProperty('ATTRIBUTE_GROUPS', JSON.stringify(groups));
  return groups;
}

/**
 * Shows the configure BOM levels dialog
 */
function showConfigureBOMLevels() {
  var html = HtmlService.createHtmlOutputFromFile('ConfigureBOMLevels')
    .setWidth(650)
    .setHeight(550)
    .setTitle('Configure BOM Hierarchy');
  SpreadsheetApp.getUi().showModalDialog(html, 'BOM Levels');
}

/**
 * Loads BOM hierarchy configuration data
 * @return {Object} Object with hierarchy and categories
 */
function loadBOMHierarchyData() {
  return {
    hierarchy: getBOMHierarchy(),
    categories: getArenaCategories()
  };
}

/**
 * Shows the item picker sidebar
 */
function showItemPicker() {
  var html = HtmlService.createHtmlOutputFromFile('ItemPicker')
    .setTitle('Arena Item Picker');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Pulls BOM from Arena and populates sheets
 */
function pullBOMFromArena() {
  var ui = SpreadsheetApp.getUi();

  if (!isAuthorized()) {
    ui.alert('Not Configured', 'Please configure your Arena API connection first.', ui.ButtonSet.OK);
    showLoginWizard();
    return;
  }

  // Prompt for parent item number
  var response = ui.prompt(
    'Pull BOM',
    'Enter the Arena item number to pull BOM from:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var itemNumber = response.getResponseText().trim();
  if (!itemNumber) {
    ui.alert('Item number is required');
    return;
  }

  try {
    var result = pullBOM(itemNumber);
    ui.alert('Success', result.message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Failed to pull BOM: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Pushes BOM to Arena
 */
function pushBOMToArena() {
  var ui = SpreadsheetApp.getUi();

  if (!isAuthorized()) {
    ui.alert('Not Configured', 'Please configure your Arena API connection first.', ui.ButtonSet.OK);
    showLoginWizard();
    return;
  }

  var response = ui.alert(
    'Push BOM',
    'This will push the current sheet layout to Arena as a BOM. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    var result = pushBOM();
    ui.alert('Success', result.message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Failed to push BOM: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Loads data for the item picker sidebar
 * @return {Object} Object with items, categories, and colors
 */
function loadItemPickerData() {
  var arenaClient = new ArenaAPIClient();

  // Fetch all items from Arena
  var rawItems = arenaClient.getAllItems(400);

  // Map Arena API response to format expected by Item Picker
  var mappedItems = rawItems.map(function(item) {
    // Handle both lowercase and capitalized property names from Arena API
    var categoryObj = item.category || item.Category || {};
    var lifecycleObj = item.lifecyclePhase || item.LifecyclePhase || {};

    return {
      guid: item.guid || item.Guid,
      number: item.number || item.Number || '',
      name: item.name || item.Name || '',
      description: item.description || item.Description || '',
      revisionNumber: item.revisionNumber || item.RevisionNumber || item.revision || item.Revision || '',
      categoryGuid: categoryObj.guid || categoryObj.Guid || '',
      categoryName: categoryObj.name || categoryObj.Name || '',
      categoryPath: categoryObj.path || categoryObj.Path || '',
      lifecyclePhase: lifecycleObj.name || lifecycleObj.Name || '',
      lifecyclePhaseGuid: lifecycleObj.guid || lifecycleObj.Guid || '',
      attributes: item.attributes || item.Attributes || []
    };
  });

  Logger.log('Mapped ' + mappedItems.length + ' items for Item Picker');

  return {
    items: mappedItems,
    categories: getCategoriesWithFavorites(),
    colors: getCategoryColors()
  };
}

/**
 * Stores the selected item from the item picker
 * @param {Object} item - Selected item data
 */
function setSelectedItem(item) {
  PropertiesService.getUserProperties().setProperty('SELECTED_ITEM', JSON.stringify(item));
}

/**
 * Gets the currently selected item
 * @return {Object} Selected item or null
 */
function getSelectedItem() {
  var json = PropertiesService.getUserProperties().getProperty('SELECTED_ITEM');
  if (!json) return null;

  try {
    return JSON.parse(json);
  } catch (e) {
    return null;
  }
}

/**
 * Clears the selected item
 */
function clearSelectedItem() {
  PropertiesService.getUserProperties().deleteProperty('SELECTED_ITEM');
}

/**
 * Gets item quantities from the current sheet
 * Counts how many times each Arena item number appears in the sheet
 * @return {Object} Map of item numbers to quantities
 */
function getItemQuantities() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var quantities = {};

    // Get all item numbers from Arena to use as a lookup
    var arenaClient = new ArenaAPIClient();
    var items = arenaClient.getAllItems(400);

    // Build a set of valid item numbers for fast lookup
    var validItemNumbers = {};
    items.forEach(function(item) {
      var itemNum = item.number || item.Number;
      if (itemNum) {
        validItemNumbers[itemNum.trim()] = true;
      }
    });

    Logger.log('Tracking ' + Object.keys(validItemNumbers).length + ' valid item numbers');

    // Skip header row (row 0)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // Look for item numbers in each cell
      for (var j = 0; j < row.length; j++) {
        var cellValue = row[j];

        // Check if cell contains a valid Arena item number
        if (cellValue && typeof cellValue === 'string') {
          var trimmed = cellValue.trim();

          // Only count if this is a valid Arena item number
          if (trimmed && validItemNumbers[trimmed]) {
            if (quantities[trimmed]) {
              quantities[trimmed]++;
            } else {
              quantities[trimmed] = 1;
            }
          }
        }
      }
    }

    Logger.log('Found ' + Object.keys(quantities).length + ' unique items in sheet');
    return quantities;

  } catch (error) {
    Logger.log('Error getting item quantities: ' + error.message);
    return {};
  }
}

/**
 * Inserts a selected item into the active cell
 * Called when user clicks a cell after selecting an item
 */
function insertSelectedItem() {
  var selectedItem = getSelectedItem();
  if (!selectedItem) {
    SpreadsheetApp.getUi().alert('No Item Selected', 'Please select an item from the Item Picker first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();

  if (!cell) {
    SpreadsheetApp.getUi().alert('No Cell Selected', 'Please select a cell to insert the item.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Insert item number into cell (replace mode)
  cell.setValue(selectedItem.number);

  // Apply category color
  var categoryColor = getCategoryColor(selectedItem.categoryName);
  if (categoryColor) {
    cell.setBackground(categoryColor);
  }

  // Populate attributes in columns to the right
  populateItemAttributes(cell.getRow(), selectedItem);

  // Clear selection
  clearSelectedItem();
}

/**
 * Populates item attributes in columns to the right of the item number
 * @param {number} row - Row number where item was inserted
 * @param {Object} item - Item data with attributes
 */
function populateItemAttributes(row, item) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var columns = getItemColumns();

  if (!columns || columns.length === 0) {
    return; // No columns configured
  }

  var startCol = sheet.getActiveCell().getColumn() + 1;

  columns.forEach(function(col, index) {
    var targetCol = startCol + index;
    var cell = sheet.getRange(row, targetCol);

    // Find the attribute value in the item
    var value = getAttributeValue(item, col.attributeGuid);
    if (value) {
      cell.setValue(value);
    }
  });
}
