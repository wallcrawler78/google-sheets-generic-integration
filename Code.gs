/**
 * PTC Arena Sheets Generic Integration
 * Main entry point for Google Sheets Add-on
 * REFACTORED: Now supports user-configurable terminology via TypeSystemConfig.gs
 */

/**
 * Runs when the add-on is installed
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Runs when the spreadsheet is opened
 * REFACTORED: Adds first-run detection and auto-migration
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();

  // Check for first run and handle initialization
  var firstRunCheck = checkFirstRun();

  if (firstRunCheck.action === 'auto-migrate') {
    // Silently auto-migrate datacenter configuration
    Logger.log('Auto-migrating datacenter configuration...');
    var migrationResult = autoMigrateIfNeeded();

    if (migrationResult.performed && migrationResult.success) {
      Logger.log('Auto-migration successful');

      // Show one-time notification if needed
      if (shouldShowMigrationNotification()) {
        showMigrationNotification();
      }
    }
  } else if (firstRunCheck.action === 'show-wizard') {
    // First run - show setup wizard to configure terminology
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Welcome! Please complete the setup wizard to configure your integration.',
      'First Time Setup',
      5
    );

    // Show wizard after a short delay to let menu load
    Utilities.sleep(500);
    showSetupWizard();
    return; // Exit early - wizard will trigger reload when complete
  }

  // Get dynamic terminology for menu labels
  var entitySingular = getTerminology('entity_singular');
  var entityPlural = getTerminology('entity_plural');
  var level0Name = getTerminology('hierarchy_level_0');

  // Create menu with dynamic terminology
  ui.createMenu('Arena ' + level0Name)
    .addSubMenu(ui.createMenu('Create Layout')
      .addItem('New ' + entitySingular + ' Configuration', 'createNewRackConfiguration')
      .addItem('New Overview Layout', 'createNewOverviewLayout')
      .addItem('Autolink ' + entityPlural + ' to Overview', 'autoLinkRacksToOverviewAction'))
    .addSeparator()
    .addItem('Show Item Picker', 'showItemPicker')
    .addItem('Show ' + entitySingular + ' Picker', 'showRackPicker')
    .addSeparator()
    .addSubMenu(ui.createMenu('BOM Operations')
      .addItem('Create Consolidated BOM', 'createConsolidatedBOMSheet')
      .addItem('Push ' + level0Name + ' Structure to Arena (NEW)', 'pushPODStructureToArenaNew')
      .addItem('Push ' + level0Name + ' Structure to Arena (OLD)', 'pushPODStructureToArena')
      .addSeparator()
      .addItem('Check all ' + entitySingular + ' Status (vs. PDM)', 'checkAllRackStatuses')
      .addItem('View ' + entitySingular + ' History', 'viewRackHistory'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Setup')
      .addSubMenu(ui.createMenu('Configuration')
        .addItem('Configure Arena Connection', 'showLoginWizard')
        .addItem('Test Connection', 'testArenaConnection')
        .addItem('Clear Credentials', 'clearCredentials')
        .addSeparator()
        .addItem('Configure Type System', 'showConfigureTypeSystem')
        .addItem('Quick Setup (Default Config)', 'showQuickSetup')
        .addItem('Run Auto-Migration (Datacenter)', 'runManualMigration')
        .addSeparator()
        .addItem('Configure Item Columns', 'showConfigureColumns')
        .addItem('Configure Category Colors', 'showConfigureColors')
        .addItem('Configure ' + entitySingular + ' Colors', 'showConfigureRackColors')
        .addItem('Configure BOM Levels', 'showConfigureBOMLevels')
        .addItem(entitySingular + ' Location Propagation', 'showRackBOMLocationSetting'))
      .addSubMenu(ui.createMenu('Advanced')
        .addItem('Mark Current ' + entitySingular + ' as Synced', 'markCurrentRackAsSynced')
        .addItem('Validate History Integrity', 'validateHistoryTabIntegrity')
        .addItem('Repair History Issues', 'repairHistoryTabIntegrity')
        .addSeparator()
        .addItem('Export Configuration', 'showExportDialog')
        .addItem('Import Configuration', 'showImportDialog')
        .addItem('Reset Configuration', 'resetConfigurationDialog')))
    .addSeparator()
    .addItem('Help and Documentation', 'showHelp')
    .addToUi();
}

/**
 * Runs when selection changes (e.g., clicking on a cell)
 * Auto-opens History Filter sidebar when navigating to Rack History tab
 */
function onSelectionChange(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();

    // Auto-open sidebar when History tab is activated
    if (sheet && sheet.getName() === 'Rack History') {
      Logger.log('Rack History tab detected - opening sidebar');

      // Check if sidebar already shown this session
      var cache = CacheService.getScriptCache();
      var shown = cache.get('history_sidebar_shown');

      Logger.log('Cache check - shown: ' + shown);

      if (!shown) {
        Logger.log('Opening sidebar for first time this session');
        showHistoryFilterSidebar();
        // Set flag for 6 hours (cache expires after that)
        cache.put('history_sidebar_shown', 'true', 21600);
      } else {
        Logger.log('Sidebar already shown this session - skipping auto-open');
      }
    }
  } catch (error) {
    Logger.log('Error in onSelectionChange: ' + error.message);
    Logger.log('Error stack: ' + error.stack);
  }
}

/**
 * Simple trigger that runs whenever a cell is edited
 * Detects local changes to rack BOMs and updates status to LOCAL_MODIFIED
 */
function onEdit(e) {
  try {
    // Only process if we have an event object
    if (!e || !e.range) return;

    var sheet = e.range.getSheet();
    var editedRow = e.range.getRow();

    // Only monitor rack config sheets
    if (!isRackConfigSheet(sheet)) return;

    // Only monitor BOM data rows (row 3+), not metadata or headers
    if (editedRow < DATA_START_ROW) return;

    // Check current status - only update if currently SYNCED
    var currentStatus = getRackSheetStatus(sheet);
    if (!currentStatus || currentStatus !== RACK_STATUS.SYNCED) {
      return; // Already marked as modified or placeholder, no need to update
    }

    // Detect if BOM has changed from synced state
    var hasLocalChanges = detectLocalChanges(sheet);

    if (hasLocalChanges) {
      // Get Arena GUID for status update
      var arenaGuid = sheet.getRange(METADATA_ROW, META_ARENA_GUID_COL).getValue();

      // Update status to LOCAL_MODIFIED
      Logger.log('Local edit detected in rack: ' + sheet.getName() + ' - updating status to LOCAL_MODIFIED');
      updateRackSheetStatus(sheet, RACK_STATUS.LOCAL_MODIFIED, arenaGuid || null);
    }

  } catch (error) {
    // Silent failure - don't interrupt user's editing experience
    Logger.log('Error in onEdit trigger: ' + error.message);
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
      // Fun welcome messages - pick one at random
      var welcomeMessages = [
        'Welcome to the Arena Data Center!',
        'Connection successful! Ready to manage your data!',
        'All systems go! Your Arena is ready!',
        'Connected! Time to build something awesome!',
        'Success! Your data center awaits!',
        'Arena connection established! Let\'s get to work!'
      ];
      var randomWelcome = welcomeMessages[Math.floor(Math.random() * welcomeMessages.length)];

      var message = randomWelcome + '\n\n' +
                   'Workspace: ' + getWorkspaceId() + '\n' +
                   'Total Items: ' + result.totalItems + '\n' +
                   'Categories: ' + result.categoryCount + '\n\n' +
                   'Ready to create rack configurations and BOMs!';

      ui.alert('Connection Successful', message, ui.ButtonSet.OK);
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
    var groupedItems = groupItemsByEntityType(arenaItems); // REFACTORED: Use generic entity type

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
 * Shows the rack color configuration dialog
 */
function showConfigureRackColors() {
  var html = HtmlService.createHtmlOutputFromFile('ConfigureRackColors')
    .setWidth(650)
    .setHeight(600)
    .setTitle('Configure Rack Colors');
  SpreadsheetApp.getUi().showModalDialog(html, 'Rack Colors');
}

/**
 * Loads category and color data for the configuration dialog
 * @return {Object} Object with categories and colors
 */
function loadCategoryColorData() {
  return {
    categories: getCategoriesWithFavorites(),
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
    categories: getCategoriesWithFavorites()
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
 * Shows the rack picker sidebar
 * REFACTORED: Now uses dynamic terminology
 */
function showRackPicker() {
  var entitySingular = getTerminology('entity_singular');
  var html = HtmlService.createHtmlOutputFromFile('RackPicker')
    .setTitle(entitySingular + ' Picker');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Gets terminology for client-side UI
 * Called from HTML files to get dynamic terminology
 * @return {Object} Terminology object with all keys
 */
function getTerminologyForUI() {
  return {
    entity_singular: getTerminology('entity_singular'),
    entity_plural: getTerminology('entity_plural'),
    entity_verb: getTerminology('entity_verb'),
    entity_singular_lower: getTerminology('entity_singular_lower'),
    entity_plural_lower: getTerminology('entity_plural_lower'),
    hierarchy_level_0: getTerminology('hierarchy_level_0'),
    hierarchy_level_1: getTerminology('hierarchy_level_1'),
    hierarchy_level_2: getTerminology('hierarchy_level_2')
  };
}

/**
 * Shows the help modal with documentation and guides
 */
function showHelp() {
  var html = HtmlService.createHtmlOutputFromFile('HelpModal')
    .setWidth(850)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Arena Data Center - Help Guide');
}

/**
 * Shows category selector dialog
 * @param {string} title - Dialog title
 * @param {string} subtitle - Dialog subtitle
 * @return {Object} Selected category {guid, name} or null if cancelled
 */
function showCategorySelector(title, subtitle) {
  Logger.log('=== CATEGORY SELECTOR START ===');
  Logger.log('Title: ' + title);
  Logger.log('Subtitle: ' + subtitle);

  // Store dialog parameters
  PropertiesService.getUserProperties().setProperty('category_selector_title', title);
  PropertiesService.getUserProperties().setProperty('category_selector_subtitle', subtitle);
  PropertiesService.getUserProperties().deleteProperty('category_selection');

  var html = HtmlService.createHtmlOutputFromFile('CategorySelector')
    .setWidth(550)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, title);
  Logger.log('Category selector dialog displayed, waiting for user selection...');

  // Poll for selection with timeout
  // NOTE: This is a workaround for Apps Script's lack of true blocking dialogs
  // The dialog is shown asynchronously, so we poll every second for up to 5 minutes
  var maxAttempts = 300; // 5 minutes (300 seconds)
  var attempts = 0;
  var selection = null;

  while (attempts < maxAttempts) {
    Utilities.sleep(1000); // Wait 1 second
    selection = PropertiesService.getUserProperties().getProperty('category_selection');

    if (selection) {
      Logger.log('Category selected after ' + attempts + ' seconds');
      Logger.log('Raw selection JSON: ' + selection);
      var parsedSelection = JSON.parse(selection);
      Logger.log('Parsed selection object: ' + JSON.stringify(parsedSelection));
      Logger.log('Selected category GUID: ' + parsedSelection.guid);
      Logger.log('Selected category name: ' + parsedSelection.name);
      Logger.log('=== CATEGORY SELECTOR END (SUCCESS) ===');
      return parsedSelection;
    }

    attempts++;
  }

  // Timeout - user didn't make a selection
  Logger.log('Category selection timed out after ' + maxAttempts + ' seconds');
  Logger.log('=== CATEGORY SELECTOR END (TIMEOUT) ===');
  return null;
}

/**
 * Gets data for category selector dialog
 * @return {Object} Data object with categories, favorites, title, subtitle
 */
function getCategorySelectorData() {
  var categories = getArenaCategories();
  var favoriteGuids = getFavoriteCategories();  // This returns category GUIDs

  // Get favorite category objects by matching GUIDs
  var favoriteCategoryObjects = [];
  if (favoriteGuids && favoriteGuids.length > 0) {
    favoriteCategoryObjects = categories.filter(function(cat) {
      return favoriteGuids.indexOf(cat.guid) !== -1;
    });
  }

  Logger.log('Found ' + favoriteCategoryObjects.length + ' favorite categories out of ' + categories.length + ' total');

  return {
    categories: categories,
    favorites: favoriteCategoryObjects,
    title: PropertiesService.getUserProperties().getProperty('category_selector_title') || 'Select Category',
    subtitle: PropertiesService.getUserProperties().getProperty('category_selector_subtitle') || ''
  };
}

/**
 * Sets the category selection from dialog
 * @param {Object} category - Selected category {guid, name} or null
 */
function setCategorySelection(category) {
  Logger.log('=== SET CATEGORY SELECTION ===');
  if (category) {
    Logger.log('Category received from dialog: ' + JSON.stringify(category));
    Logger.log('Category GUID: ' + category.guid);
    Logger.log('Category name: ' + category.name);
    PropertiesService.getUserProperties().setProperty('category_selection', JSON.stringify(category));
    Logger.log('Category selection saved to user properties');
  } else {
    Logger.log('Category selection cleared (user cancelled)');
    PropertiesService.getUserProperties().deleteProperty('category_selection');
  }

  // Return the category so caller can use it
  return category;
}

/**
 * Gets the category selection from user properties
 * Used by wizard to retrieve selection after category selector closes
 * @return {Object|null} Selected category {guid, name} or null
 */
function getCategorySelection() {
  var json = PropertiesService.getUserProperties().getProperty('category_selection');
  if (!json) return null;

  try {
    return JSON.parse(json);
  } catch (error) {
    Logger.log('Error parsing category selection: ' + error.message);
    return null;
  }
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
function loadItemPickerData(forceRefresh) {
  try {
    var cache = CacheService.getScriptCache();
    var itemsCacheKey = 'itemPicker_items_v2';
    var categoriesCacheKey = 'itemPicker_categories_v2';
    var colorsCacheKey = 'itemPicker_colors_v2';

    // Try to get from cache first (unless force refresh)
    if (!forceRefresh) {
      var cachedItems = cache.get(itemsCacheKey);
      var cachedCategories = cache.get(categoriesCacheKey);
      var cachedColors = cache.get(colorsCacheKey);

      if (cachedItems && cachedCategories) {
        Logger.log('Loading Item Picker data from cache');
        try {
          var items = JSON.parse(cachedItems);
          var categories = JSON.parse(cachedCategories);
          var colors = cachedColors ? JSON.parse(cachedColors) : {};
          Logger.log('✅ Cache hit! Loaded ' + items.length + ' items, ' + categories.length + ' categories from cache');
          return { items: items, categories: categories, colors: colors };
        } catch (parseError) {
          Logger.log('⚠️ Cache data corrupted, fetching fresh data: ' + parseError.message);
        }
      } else {
        Logger.log('ℹ️ No cache found, fetching from Arena');
      }
    } else {
      Logger.log('🔄 Force refresh requested, bypassing cache');
    }

    // Cache miss or force refresh - fetch from Arena
    var arenaClient = new ArenaAPIClient();

    // Fetch all items from Arena
    Logger.log('Fetching items from Arena...');
    var rawItems = arenaClient.getAllItems(400);
    Logger.log('Received ' + rawItems.length + ' raw items from Arena');

    // Map Arena API response to format expected by Item Picker
    var mappedItems = rawItems.map(function(item) {
      // Handle both lowercase and capitalized property names from Arena API
      var categoryObj = item.category || item.Category || {};
      var lifecycleObj = item.lifecyclePhase || item.LifecyclePhase || {};

      // Log first item's raw structure to understand Arena's field names
      if (rawItems.indexOf(item) === 0) {
        Logger.log('=== RAW ARENA ITEM STRUCTURE (responseview=full) ===');
        Logger.log('Available fields: ' + Object.keys(item).join(', '));
        Logger.log('description field: ' + (item.description || item.Description || 'NOT FOUND'));

        // Log badge-related fields specifically
        Logger.log('--- BADGE DETECTION FIELDS ---');
        Logger.log('modifiedFiles field: ' + (item.modifiedFiles !== undefined ? item.modifiedFiles : 'NOT FOUND'));
        Logger.log('futureChanges field: ' + (item.futureChanges ? JSON.stringify(item.futureChanges) : 'NOT FOUND'));
        Logger.log('FutureChanges field: ' + (item.FutureChanges ? JSON.stringify(item.FutureChanges) : 'NOT FOUND'));

        // Log URL fields
        Logger.log('--- URL FIELDS ---');
        var urlObj = item.url || item.Url || {};
        Logger.log('url.api: ' + (urlObj.api || urlObj.Api || 'NOT FOUND'));
        Logger.log('url.app: ' + (urlObj.app || urlObj.App || 'NOT FOUND'));

        Logger.log('--- FULL ITEM OBJECT ---');
        Logger.log(JSON.stringify(item, null, 2));
        Logger.log('=== END RAW ARENA ITEM STRUCTURE ===');
      }

      var itemNumber = item.number || item.Number || '';

      // Use Arena's provided web URL (item.url.app)
      var urlObj = item.url || item.Url || {};
      var arenaWebURL = urlObj.app || urlObj.App || '';

      // Fallback to search URL if Arena doesn't provide url.app
      if (!arenaWebURL) {
        arenaWebURL = 'https://app.bom.com/search?query=' + encodeURIComponent(itemNumber);
      }

      // Check for pending changes (ECO)
      var hasPendingChanges = false;
      var hasFiles = false;  // Not using file badges - removed for cleaner UI

      // Check for pending changes (futureChanges not in response, will need Phase 2)
      if (item.futureChanges || item.FutureChanges) {
        var futureChanges = item.futureChanges || item.FutureChanges || [];
        hasPendingChanges = Array.isArray(futureChanges) && futureChanges.length > 0;
      }

      // Only store fields needed for Item Picker display (reduces cache size)
      return {
        guid: item.guid || item.Guid,
        number: itemNumber,
        name: item.name || item.Name || '',
        description: item.description || item.Description || '',
        revisionNumber: item.revisionNumber || item.RevisionNumber || item.revision || item.Revision || '',
        categoryGuid: categoryObj.guid || categoryObj.Guid || '',
        categoryName: categoryObj.name || categoryObj.Name || '',
        lifecyclePhase: lifecycleObj.name || lifecycleObj.Name || '',
        arenaWebURL: arenaWebURL
      };
    });

    Logger.log('Mapped ' + mappedItems.length + ' items for Item Picker');

    // Log badge detection summary
    var itemsWithECO = mappedItems.filter(function(i) { return i.hasPendingChanges; }).length;
    var itemsWithFiles = mappedItems.filter(function(i) { return i.hasFiles; }).length;
    Logger.log('=== BADGE DETECTION SUMMARY ===');
    Logger.log('Items with pending ECOs: ' + itemsWithECO + ' / ' + mappedItems.length);
    Logger.log('Items with files: ' + itemsWithFiles + ' / ' + mappedItems.length);
    Logger.log('=== END BADGE SUMMARY ===');

    // Log first item for debugging
    if (mappedItems.length > 0) {
      Logger.log('Sample mapped item: ' + JSON.stringify(mappedItems[0]));
    }

    var categories = getCategoriesWithFavorites();
    var colors = getCategoryColors();

    Logger.log('Loaded ' + categories.length + ' categories');

    var result = {
      items: mappedItems,
      categories: categories,
      colors: colors
    };

    // Cache items, categories, and colors separately to stay under 100KB per key limit
    // Use 6-hour TTL to match ArenaAPIClient cache
    var cacheTTL = 6 * 60 * 60; // 6 hours in seconds

    try {
      // Cache items (largest data set)
      var itemsJson = JSON.stringify(mappedItems);
      cache.put(itemsCacheKey, itemsJson, cacheTTL);
      Logger.log('💾 Cached ' + mappedItems.length + ' items (' + Math.round(itemsJson.length / 1024) + ' KB)');
    } catch (itemsError) {
      Logger.log('⚠️ Could not cache items: ' + itemsError.message);
    }

    try {
      // Cache categories
      var categoriesJson = JSON.stringify(categories);
      cache.put(categoriesCacheKey, categoriesJson, cacheTTL);
      Logger.log('💾 Cached ' + categories.length + ' categories (' + Math.round(categoriesJson.length / 1024) + ' KB)');
    } catch (categoriesError) {
      Logger.log('⚠️ Could not cache categories: ' + categoriesError.message);
    }

    try {
      // Cache colors
      var colorsJson = JSON.stringify(colors);
      cache.put(colorsCacheKey, colorsJson, cacheTTL);
      Logger.log('💾 Cached colors (' + Math.round(colorsJson.length / 1024) + ' KB)');
    } catch (colorsError) {
      Logger.log('⚠️ Could not cache colors: ' + colorsError.message);
    }

    return result;
  } catch (error) {
    Logger.log('ERROR in loadItemPickerData: ' + error.message + '\n' + error.stack);
    throw error;
  }
}

/**
 * Stores the selected item from the item picker
 * @param {Object} item - Selected item data
 */
function setSelectedItem(item) {
  Logger.log('=== SET SELECTED ITEM ===');
  Logger.log('Item: ' + JSON.stringify(item));
  Logger.log('Item Number: ' + (item ? item.number : 'null'));
  Logger.log('Item Name: ' + (item ? item.name : 'null'));
  Logger.log('Item GUID: ' + (item ? item.guid : 'null'));

  PropertiesService.getUserProperties().setProperty('SELECTED_ITEM', JSON.stringify(item));
  Logger.log('Item stored in user properties');
  Logger.log('=== SET SELECTED ITEM COMPLETE ===');
}

/**
 * Gets the currently selected item
 * @return {Object} Selected item or null
 */
function getSelectedItem() {
  var json = PropertiesService.getUserProperties().getProperty('SELECTED_ITEM');
  if (!json) {
    Logger.log('getSelectedItem: No item in storage');
    return null;
  }

  try {
    var item = JSON.parse(json);
    Logger.log('getSelectedItem: Retrieved item: ' + item.number);
    Logger.log('getSelectedItem: Item has name: ' + (item.name ? 'YES' : 'NO'));
    Logger.log('getSelectedItem: Item has description: ' + (item.description ? 'YES' : 'NO'));
    Logger.log('getSelectedItem: Item has categoryName: ' + (item.categoryName ? 'YES' : 'NO'));
    return item;
  } catch (e) {
    Logger.log('getSelectedItem: Error parsing JSON: ' + e.message);
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
 * Checks if the active sheet is a rack configuration sheet
 * @return {boolean} True if active sheet is a rack config
 */
function isActiveSheetRackConfig() {
  var sheet = SpreadsheetApp.getActiveSheet();
  return isRackConfigSheet(sheet);
}

/**
 * Refreshes the BOM for the currently active rack sheet
 * Compares with Arena data and highlights changes
 * @return {Object} Result with success status and message
 */
function refreshCurrentRackBOM() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();

  Logger.log('=== REFRESH RACK BOM ===');

  // Verify this is a rack config sheet
  if (!isRackConfigSheet(sheet)) {
    return {
      success: false,
      message: 'Active sheet is not a rack configuration. Please open a rack sheet first.'
    };
  }

  // Get rack metadata
  var metadata = getRackConfigMetadata(sheet);
  Logger.log('Refreshing BOM for rack: ' + metadata.itemNumber);

  // Get Arena GUID from History tab
  var arenaGuid = getRackArenaGuidFromHistory(metadata.itemNumber);

  if (!arenaGuid) {
    return {
      success: false,
      message: 'This rack has no Arena GUID. It may be a placeholder not yet pushed to Arena.'
    };
  }

  // Inform user about the refresh process
  ui.alert(
    'Refreshing BOM from Arena',
    'This will:\n' +
    '• Fetch the latest BOM from Arena for rack ' + metadata.itemNumber + '\n' +
    '• Compare with your current sheet data\n' +
    '• Show any differences detected\n\n' +
    '⏱️ This may take 10-20 seconds depending on BOM size.\n\n' +
    'Click OK to continue...',
    ui.ButtonSet.OK
  );

  try {
    // Fetch current BOM from Arena
    var arenaClient = new ArenaAPIClient();
    var bomResponse = arenaClient.makeRequest('/items/' + arenaGuid + '/bom', { method: 'GET' });
    var arenaBOM = bomResponse.results || bomResponse.Results || [];

    if (!arenaBOM || !arenaBOM.length) {
      return {
        success: false,
        message: 'No BOM found in Arena for item: ' + metadata.itemNumber
      };
    }

    Logger.log('Fetched ' + arenaBOM.length + ' BOM lines from Arena');

    // Get current sheet data
    var currentData = getCurrentRackBOMData(sheet);
    Logger.log('Current sheet has ' + currentData.length + ' rows');

    // Compare and detect changes (pass arenaClient to fetch full item details)
    var changes = compareBOMs(currentData, arenaBOM, arenaClient);

    Logger.log('Changes detected:');
    Logger.log('- Modified: ' + changes.modified.length);
    Logger.log('- Added: ' + changes.added.length);
    Logger.log('- Removed: ' + changes.removed.length);

    // If no changes, return success
    if (changes.modified.length === 0 && changes.added.length === 0 && changes.removed.length === 0) {
      // Update last refreshed timestamp
      updateLastRefreshedTimestamp(sheet);

      // Update status to SYNCED - no differences with Arena
      Logger.log('No changes detected - updating status to SYNCED');
      var eventDetails = {
        changesSummary: 'No changes detected',
        details: 'Refreshed from Arena - BOM matches exactly'
      };
      updateRackSheetStatus(sheet, RACK_STATUS.SYNCED, arenaGuid, eventDetails);

      // Log refresh event
      addRackHistoryEvent(metadata.itemNumber, HISTORY_EVENT.REFRESH_NO_CHANGES, {
        details: 'Refreshed BOM from Arena - no changes detected',
        statusAfter: RACK_STATUS.SYNCED
      });

      return {
        success: true,
        message: 'BOM is up to date! No changes detected.\n\nLast refreshed: ' + new Date().toLocaleString()
      };
    }

    // Build enhanced change summary with Arena link
    var arenaClient = new ArenaAPIClient();
    var rackArenaItem = arenaClient.getItemByNumber(metadata.itemNumber);
    var arenaUrl = buildArenaItemURLFromItem(rackArenaItem, metadata.itemNumber);

    var totalChanges = changes.modified.length + changes.added.length + changes.removed.length;
    var changeMessage = 'Arena has ' + totalChanges + ' change(s) for this rack:\n\n';

    changeMessage += '🏷️ Rack: ' + metadata.itemName + '\n';
    changeMessage += '   Item #: ' + metadata.itemNumber + '\n\n';

    // Summarize changes
    if (changes.modified.length > 0) {
      changeMessage += '✏️ Modified: ' + changes.modified.length + ' item(s)\n';
      changes.modified.slice(0, 2).forEach(function(mod) {
        var fieldList = mod.changes.map(function(c) { return c.field; }).join(', ');
        changeMessage += '   • ' + mod.itemNumber + ': ' + fieldList + '\n';
      });
      if (changes.modified.length > 2) {
        changeMessage += '   ... and ' + (changes.modified.length - 2) + ' more\n';
      }
    }

    if (changes.added.length > 0) {
      changeMessage += '➕ Added: ' + changes.added.length + ' item(s)\n';
      changes.added.slice(0, 2).forEach(function(add) {
        changeMessage += '   • ' + add.itemNumber + ' - ' + add.name + '\n';
      });
      if (changes.added.length > 2) {
        changeMessage += '   ... and ' + (changes.added.length - 2) + ' more\n';
      }
    }

    if (changes.removed.length > 0) {
      changeMessage += '➖ Removed: ' + changes.removed.length + ' item(s)\n';
      changes.removed.slice(0, 2).forEach(function(rem) {
        changeMessage += '   • ' + rem.itemNumber + ' - ' + rem.name + '\n';
      });
      if (changes.removed.length > 2) {
        changeMessage += '   ... and ' + (changes.removed.length - 2) + ' more\n';
      }
    }

    changeMessage += '\n🔗 Compare in Arena:\n' + arenaUrl + '\n\n';
    changeMessage += '⚠️ Applying changes will overwrite your local rack data.\n\n';
    changeMessage += 'Apply these changes from Arena?';

    var response = ui.alert(
      'BOM Changes Detected',
      changeMessage,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      // User declined to apply changes - mark as out of sync
      Logger.log('User declined to apply changes - updating status to ARENA_MODIFIED');
      var changesCount = changes.modified.length + changes.added.length + changes.removed.length;
      var eventDetails = {
        changesSummary: changesCount + ' changes found, user declined',
        details: 'User chose not to apply Arena BOM changes'
      };
      updateRackSheetStatus(sheet, RACK_STATUS.ARENA_MODIFIED, arenaGuid, eventDetails);

      // Log refresh declined event
      addRackHistoryEvent(metadata.itemNumber, HISTORY_EVENT.REFRESH_DECLINED, {
        changesSummary: changesCount + ' changes pending',
        details: 'Modified: ' + changes.modified.length + ', Added: ' + changes.added.length + ', Removed: ' + changes.removed.length,
        statusAfter: RACK_STATUS.ARENA_MODIFIED
      });

      return {
        success: false,
        message: 'Refresh cancelled by user. No changes were applied.'
      };
    }

    // Apply changes
    applyBOMChanges(sheet, changes, metadata.itemNumber);

    // Update last refreshed timestamp
    updateLastRefreshedTimestamp(sheet);

    // Update status to SYNCED - changes applied, now matches Arena
    Logger.log('Changes applied successfully - updating status to SYNCED');
    var changesCount = changes.modified.length + changes.added.length + changes.removed.length;
    var eventDetails = {
      changesSummary: changesCount + ' changes applied',
      details: 'Modified: ' + changes.modified.length + ', Added: ' + changes.added.length + ', Removed: ' + changes.removed.length
    };
    updateRackSheetStatus(sheet, RACK_STATUS.SYNCED, arenaGuid, eventDetails);

    // Log refresh accepted event
    addRackHistoryEvent(metadata.itemNumber, HISTORY_EVENT.REFRESH_ACCEPTED, {
      changesSummary: changesCount + ' changes applied from Arena',
      details: 'Modified: ' + changes.modified.length + ', Added: ' + changes.added.length + ', Removed: ' + changes.removed.length,
      statusAfter: RACK_STATUS.SYNCED,
      link: '=HYPERLINK("#BOM History", "View Details")'
    });

    // Return success message
    var summary = (changes.modified.length + changes.added.length + changes.removed.length) + ' changes applied';
    if (changes.modified.length > 0) summary += '\n• ' + changes.modified.length + ' items updated (red text)';
    if (changes.added.length > 0) summary += '\n• ' + changes.added.length + ' items added';
    if (changes.removed.length > 0) summary += '\n• ' + changes.removed.length + ' items removed';
    summary += '\n\nCheck "BOM History" tab for details.';

    return {
      success: true,
      message: summary
    };

  } catch (error) {
    Logger.log('ERROR in refreshCurrentRackBOM: ' + error.message);
    Logger.log(error.stack);

    // Update status to ERROR if sync failed
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      if (isRackConfigSheet(sheet)) {
        var metadata = getRackConfigMetadata(sheet);
        var arenaGuid = getRackArenaGuidFromHistory(metadata.itemNumber);

        var eventDetails = {
          changesSummary: 'Error during refresh',
          details: 'Error: ' + error.message
        };
        updateRackSheetStatus(sheet, RACK_STATUS.ERROR, arenaGuid || null, eventDetails);

        // Log error event
        addRackHistoryEvent(metadata.itemNumber, HISTORY_EVENT.ERROR, {
          changesSummary: 'Refresh failed',
          details: error.message,
          statusAfter: RACK_STATUS.ERROR
        });
      }
    } catch (statusError) {
      Logger.log('Failed to update error status: ' + statusError.message);
    }

    return {
      success: false,
      message: 'Error refreshing BOM: ' + error.message
    };
  }
}

/**
 * Gets item quantities from the current sheet
 * Counts how many times each Arena item number appears in the sheet
 * @return {Object} Map of item numbers to quantities
 */
function getItemQuantities() {
  return getItemQuantitiesWithScope('current').quantities;
}

/**
 * Gets item quantities with scope filtering
 * @param {string} scope - Scope: 'current', 'overview', 'racks', or 'all'
 * @return {Object} Object with quantities map and metadata
 */
function getItemQuantitiesWithScope(scope) {
  try {
    scope = scope || 'current';
    Logger.log('Getting item quantities with scope: ' + scope);

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var quantities = {};
    var sheetsToScan = [];

    // Determine which sheets to scan based on scope
    if (scope === 'current') {
      sheetsToScan = [SpreadsheetApp.getActiveSheet()];
    } else if (scope === 'overview') {
      // Find overview sheets
      var allSheets = spreadsheet.getSheets();
      allSheets.forEach(function(sheet) {
        if (sheet.getName().toLowerCase().indexOf('overview') !== -1) {
          sheetsToScan.push(sheet);
        }
      });
    } else if (scope === 'racks') {
      // For "All Racks" scope, we need to aggregate properly:
      // 1. Find overview sheet
      // 2. Count how many times each rack is placed
      // 3. For each rack, get its BOM and multiply quantities

      Logger.log('Using rack aggregation logic for accurate quantities');

      // Find overview sheet
      var allSheets = spreadsheet.getSheets();
      var overviewSheet = null;
      for (var i = 0; i < allSheets.length; i++) {
        if (allSheets[i].getName().toLowerCase().indexOf('overview') !== -1) {
          overviewSheet = allSheets[i];
          break;
        }
      }

      if (overviewSheet) {
        // Use the proper BOM aggregation logic
        var bomData = buildConsolidatedBOMFromOverview(overviewSheet);

        if (bomData && bomData.lines) {
          bomData.lines.forEach(function(line) {
            quantities[line.itemNumber] = line.quantity;
          });
        }

        Logger.log('Aggregated quantities from overview BOM: ' + Object.keys(quantities).length + ' items');

        return {
          quantities: quantities,
          sheetsScanned: 1,
          scope: scope
        };
      } else {
        // Fallback: just scan rack sheets without aggregation
        Logger.log('No overview sheet found, falling back to simple rack scan');
        var rackConfigs = getAllRackConfigTabs();
        rackConfigs.forEach(function(config) {
          sheetsToScan.push(config.sheet);
        });
      }
    } else if (scope === 'all') {
      sheetsToScan = spreadsheet.getSheets();
    }

    Logger.log('Scanning ' + sheetsToScan.length + ' sheets');

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

    // Scan each sheet
    sheetsToScan.forEach(function(sheet) {
      var data = sheet.getDataRange().getValues();

      // Skip header rows (row 0-2)
      for (var i = 2; i < data.length; i++) {
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
    });

    Logger.log('Found ' + Object.keys(quantities).length + ' unique items');

    return {
      quantities: quantities,
      sheetsScanned: sheetsToScan.length,
      scope: scope
    };

  } catch (error) {
    Logger.log('Error getting item quantities: ' + error.message);
    return {
      quantities: {},
      sheetsScanned: 0,
      scope: scope,
      error: error.message
    };
  }
}

/**
 * Gets rack quantities with scope filtering
 * Counts how many times each rack configuration appears in sheets
 * @param {string} scope - Scope: 'current', 'overview', 'racks', or 'all'
 * @return {Object} Object with quantities map and metadata
 */
function getRackQuantitiesWithScope(scope) {
  try {
    scope = scope || 'current';
    Logger.log('Getting rack quantities with scope: ' + scope);

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var quantities = {};
    var sheetsToScan = [];

    // Determine which sheets to scan based on scope
    if (scope === 'current') {
      sheetsToScan = [SpreadsheetApp.getActiveSheet()];
    } else if (scope === 'overview') {
      // Find overview sheets
      var allSheets = spreadsheet.getSheets();
      allSheets.forEach(function(sheet) {
        if (sheet.getName().toLowerCase().indexOf('overview') !== -1) {
          sheetsToScan.push(sheet);
        }
      });
    } else if (scope === 'racks') {
      // Find all rack config sheets
      var rackConfigs = getAllRackConfigTabs();
      rackConfigs.forEach(function(config) {
        sheetsToScan.push(config.sheet);
      });
    } else if (scope === 'all') {
      sheetsToScan = spreadsheet.getSheets();
    }

    Logger.log('Scanning ' + sheetsToScan.length + ' sheets for racks');

    // Get all valid rack item numbers from rack config tabs
    var validRackNumbers = {};
    var allRackConfigs = getAllRackConfigTabs();
    allRackConfigs.forEach(function(config) {
      validRackNumbers[config.itemNumber] = true;
    });

    Logger.log('Tracking ' + Object.keys(validRackNumbers).length + ' valid rack numbers');

    // Scan each sheet
    sheetsToScan.forEach(function(sheet) {
      var data = sheet.getDataRange().getValues();

      // Skip header rows (row 0-2)
      for (var i = 2; i < data.length; i++) {
        var row = data[i];

        // Look for rack numbers in each cell
        for (var j = 0; j < row.length; j++) {
          var cellValue = row[j];

          // Check if cell contains a valid rack configuration number
          if (cellValue && typeof cellValue === 'string') {
            var trimmed = cellValue.trim();

            // Only count if this is a valid rack configuration number
            if (trimmed && validRackNumbers[trimmed]) {
              if (quantities[trimmed]) {
                quantities[trimmed]++;
              } else {
                quantities[trimmed] = 1;
              }
            }
          }
        }
      }
    });

    Logger.log('Found ' + Object.keys(quantities).length + ' unique racks');

    return {
      quantities: quantities,
      sheetsScanned: sheetsToScan.length,
      scope: scope
    };

  } catch (error) {
    Logger.log('Error getting rack quantities: ' + error.message);
    return {
      quantities: {},
      sheetsScanned: 0,
      scope: scope,
      error: error.message
    };
  }
}

/**
 * Inserts a selected item into the active cell
 * Called when user clicks a cell after selecting an item
 */
function insertSelectedItem() {
  Logger.log('=== INSERT SELECTED ITEM ===');

  var selectedItem = getSelectedItem();
  if (!selectedItem) {
    Logger.log('ERROR: No item selected');
    SpreadsheetApp.getUi().alert('No Item Selected', 'Please select an item from the Item Picker first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  Logger.log('Selected item: ' + selectedItem.number);

  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();

  if (!cell) {
    Logger.log('ERROR: No cell selected');
    SpreadsheetApp.getUi().alert('No Cell Selected', 'Please select a cell to insert the item.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  Logger.log('Target sheet: ' + sheet.getName());
  Logger.log('Target cell: ' + cell.getA1Notation());

  // Check if this is a rack config sheet
  var isRackConfig = isRackConfigSheet(sheet);
  Logger.log('Is rack config sheet: ' + isRackConfig);

  if (isRackConfig) {
    // Insert into rack config sheet (full row with attributes + qty)
    Logger.log('Inserting into rack config at row ' + cell.getRow());
    insertItemIntoRackConfig(sheet, cell.getRow(), selectedItem);
    Logger.log('Rack config insertion complete');
  } else {
    // Insert into regular sheet (item number only, with potential hyperlink)
    Logger.log('Inserting into regular sheet');
    cell.setValue(selectedItem.number);

    // Apply category color
    var categoryColor = getCategoryColor(selectedItem.categoryName);
    if (categoryColor) {
      Logger.log('Applying category color: ' + categoryColor);
      cell.setBackground(categoryColor);
    }

    // Try to create hyperlink to rack config if this is a rack item
    var rackConfigSheet = findRackConfigTab(selectedItem.number);
    if (rackConfigSheet) {
      var sheetId = rackConfigSheet.getSheetId();
      var formula = '=HYPERLINK("#gid=' + sheetId + '", "' + selectedItem.number + '")';
      Logger.log('Creating hyperlink formula: ' + formula);
      cell.setFormula(formula);
    }
  }

  // Clear selection
  Logger.log('Clearing selected item');
  clearSelectedItem();
  Logger.log('=== INSERT SELECTED ITEM COMPLETE ===');
}

/**
 * Inserts an item into a rack configuration sheet with full details
 * @param {Sheet} sheet - Rack config sheet
 * @param {number} row - Row to insert into
 * @param {Object} item - Item data
 */
function insertItemIntoRackConfig(sheet, row, item) {
  Logger.log('=== INSERT ITEM INTO RACK CONFIG ===');
  Logger.log('Sheet: ' + sheet.getName());
  Logger.log('Row: ' + row);
  Logger.log('Item object (from picker): ' + JSON.stringify(item));

  // Fetch full item details from Arena to get description and complete data
  // The /items list endpoint doesn't include descriptions, so we need to fetch individually
  var fullItem = item;
  if (item.guid) {
    try {
      Logger.log('Fetching full item details from Arena for GUID: ' + item.guid);
      var arenaClient = new ArenaAPIClient();
      var arenaItem = arenaClient.getItem(item.guid);

      if (arenaItem) {
        Logger.log('Full item fetched from Arena');

        // Merge Arena data with picker data (Arena has full details)
        fullItem = {
          guid: item.guid,
          number: item.number,
          name: arenaItem.name || arenaItem.Name || item.name,
          description: arenaItem.description || arenaItem.Description || '',
          categoryName: item.categoryName,
          categoryGuid: item.categoryGuid,
          lifecyclePhase: item.lifecyclePhase,
          attributes: arenaItem.attributes || arenaItem.Attributes || item.attributes || []
        };

        Logger.log('Full item description: ' + fullItem.description);
      }
    } catch (fetchError) {
      Logger.log('WARNING: Could not fetch full item details: ' + fetchError.message);
      Logger.log('Continuing with basic item data from picker');
    }
  }

  Logger.log('Item number: ' + fullItem.number);
  Logger.log('Item name: ' + fullItem.name);
  Logger.log('Item description: ' + fullItem.description);
  Logger.log('Item categoryName: ' + fullItem.categoryName);
  Logger.log('Item lifecyclePhase: ' + fullItem.lifecyclePhase);
  Logger.log('Item attributes: ' + (fullItem.attributes ? fullItem.attributes.length : 'none'));

  // Column structure: Item Number | Name | Description | Category | Lifecycle | Qty | ...attributes
  var col = 1;

  // Item Number
  Logger.log('Setting item number in column ' + col);
  sheet.getRange(row, col++).setValue(fullItem.number);

  // Name
  Logger.log('Setting name in column ' + col);
  sheet.getRange(row, col++).setValue(fullItem.name || '');

  // Description
  Logger.log('Setting description in column ' + col);
  sheet.getRange(row, col++).setValue(fullItem.description || '');

  // Category
  Logger.log('Setting category in column ' + col);
  sheet.getRange(row, col++).setValue(fullItem.categoryName || '');

  // Lifecycle
  Logger.log('Setting lifecycle in column ' + col);
  sheet.getRange(row, col++).setValue(fullItem.lifecyclePhase || '');

  // Qty (default to 1)
  Logger.log('Setting quantity in column ' + col);
  sheet.getRange(row, col++).setValue(1);

  // Populate configured attributes
  var columns = getItemColumns();
  Logger.log('Populating ' + columns.length + ' attribute columns');
  columns.forEach(function(column) {
    var value = getAttributeValue(fullItem, column.attributeGuid);
    if (value) {
      Logger.log('Setting attribute "' + column.header + '" in column ' + col);
      sheet.getRange(row, col).setValue(value);
    }
    col++;
  });

  // Apply category color to entire row
  var categoryColor = getCategoryColor(fullItem.categoryName);
  if (categoryColor) {
    Logger.log('Applying category color: ' + categoryColor);
    var lastCol = sheet.getLastColumn();
    sheet.getRange(row, 1, 1, lastCol).setBackground(categoryColor);
  }

  Logger.log('=== INSERT ITEM INTO RACK CONFIG COMPLETE ===');
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

/**
 * Gets all rack configurations for the Rack Picker
 * @return {Array<Object>} Array of rack configuration metadata
 */
function getAllRackConfigurationsForPicker() {
  try {
    var rackConfigs = getAllRackConfigTabs();
    Logger.log('Found ' + rackConfigs.length + ' rack configurations');
    return rackConfigs;
  } catch (error) {
    Logger.log('Error loading rack configurations: ' + error.message);
    throw error;
  }
}

/**
 * Inserts a selected rack into the active cell(s)
 * If rack config doesn't exist, creates it and pulls BOM from Arena
 * Supports multi-cell selection for bulk placement
 * @param {Object} rack - Rack metadata object (can be existing rack config or Arena item)
 */
function insertSelectedRack(rack) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();

  if (!range) {
    throw new Error('No cell(s) selected. Please select cell(s) in the overview grid.');
  }

  var itemNumber = rack.itemNumber || rack.number;

  if (!itemNumber) {
    throw new Error('Invalid rack object: missing item number');
  }

  // Check if rack config tab exists (create once, use for all cells)
  var rackConfigSheet = findRackConfigTab(itemNumber);

  // If rack config doesn't exist, create it and pull BOM from Arena
  if (!rackConfigSheet) {
    Logger.log('Rack config not found for ' + itemNumber + ', creating from Arena item...');

    try {
      // Create rack config from Arena item
      rackConfigSheet = createRackConfigFromArenaItem(rack);
      Logger.log('Created rack config sheet: ' + rackConfigSheet.getName());

    } catch (error) {
      Logger.log('Error creating rack config: ' + error.message);
      throw new Error('Failed to create rack configuration: ' + error.message);
    }
  }

  // Get all cells in the range
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var totalCells = numRows * numCols;

  Logger.log('Placing rack ' + itemNumber + ' in ' + totalCells + ' cell(s)');

  // Prepare formula or value
  var formula = null;
  if (rackConfigSheet) {
    var sheetId = rackConfigSheet.getSheetId();
    formula = '=HYPERLINK("#gid=' + sheetId + '", "' + itemNumber + '")';
  }

  // Get rack color (custom or auto-generated)
  var rackColor = getRackColor(itemNumber);
  Logger.log('Using color ' + rackColor + ' for rack ' + itemNumber);

  // Insert into each cell in the range
  for (var row = 1; row <= numRows; row++) {
    for (var col = 1; col <= numCols; col++) {
      var cell = range.getCell(row, col);

      if (formula) {
        cell.setFormula(formula);
        cell.setFontColor('#1a73e8');
      } else {
        cell.setValue(itemNumber);
      }

      // Apply rack color as background
      cell.setBackground(rackColor);

      // Center align and make bold
      cell.setHorizontalAlignment('center');
      cell.setVerticalAlignment('middle');
      cell.setFontWeight('bold');
    }
  }

  Logger.log('Inserted rack: ' + itemNumber + ' in ' + totalCells + ' cells at ' + range.getA1Notation() + ' with color ' + rackColor);
}

/**
 * Creates a rack configuration tab from an Arena item and pulls its BOM
 * @param {Object} arenaItem - Arena item object (rack)
 * @return {Sheet} Created rack config sheet
 */
function createRackConfigFromArenaItem(arenaItem) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Extract item details (handle both formats - existing rack config or Arena item)
  var itemNumber = arenaItem.itemNumber || arenaItem.number || arenaItem.Number;
  var itemName = arenaItem.itemName || arenaItem.name || arenaItem.Name || itemNumber;
  var description = arenaItem.description || arenaItem.Description || '';
  var itemGuid = arenaItem.guid || arenaItem.Guid || null;

  Logger.log('Creating rack config for: ' + itemNumber + ' (GUID: ' + itemGuid + ')');

  // Create the new sheet
  var sheetName = 'Rack - ' + itemNumber + ' (' + itemName + ')';
  var newSheet = ss.insertSheet(sheetName);

  // Set up metadata row (Row 1)
  newSheet.getRange(1, 1).setValue('PARENT_ITEM');  // METADATA_ROW, META_LABEL_COL
  newSheet.getRange(1, 2).setValue(itemNumber);      // META_ITEM_NUM_COL
  newSheet.getRange(1, 3).setValue(itemName);        // META_ITEM_NAME_COL
  newSheet.getRange(1, 4).setValue(description);     // META_ITEM_DESC_COL

  // Format metadata row
  var metaRange = newSheet.getRange(1, 1, 1, 4);
  metaRange.setBackground('#e8f0fe');
  metaRange.setFontWeight('bold');
  metaRange.setFontColor('#1967d2');

  // Set up header row (Row 2)
  var itemColumns = getItemColumns();
  var headers = ['Item Number', 'Name', 'Description', 'Category', 'Lifecycle', 'Qty'];

  // Add configured attribute columns
  itemColumns.forEach(function(col) {
    headers.push(col.header || col.attributeName);
  });

  var headerRange = newSheet.getRange(2, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // Freeze header rows
  newSheet.setFrozenRows(2);

  // Set column widths
  newSheet.setColumnWidth(1, 120);  // Item Number
  newSheet.setColumnWidth(2, 200);  // Name
  newSheet.setColumnWidth(3, 300);  // Description
  newSheet.setColumnWidth(4, 150);  // Category
  newSheet.setColumnWidth(5, 120);  // Lifecycle
  newSheet.setColumnWidth(6, 60);   // Qty

  // Try to pull BOM from Arena and populate the sheet
  try {
    Logger.log('Pulling BOM for rack: ' + itemNumber);
    pullBOMForRack(newSheet, itemNumber, itemGuid);
    Logger.log('BOM pull completed for: ' + itemNumber);
  } catch (bomError) {
    Logger.log('Could not pull BOM from Arena: ' + bomError.message);
    Logger.log('BOM Error Stack: ' + bomError.stack);

    // Add detailed error message to sheet
    newSheet.getRange(3, 1).setValue('⚠️ BOM Pull Failed: ' + bomError.message);
    newSheet.getRange(3, 1, 1, headers.length).setFontStyle('italic').setFontColor('#ea4335');
    newSheet.getRange(4, 1).setValue('→ Use Item Picker to manually add components');
    newSheet.getRange(4, 1, 1, headers.length).setFontStyle('italic').setFontColor('#666666');
  }

  // Set tab color to cascading blue
  var rackIndex = getAllRackConfigTabs().length;  // Get count including this new one
  var blueColor = getCascadingBlueColor(rackIndex - 1);  // Subtract 1 for zero-based index
  newSheet.setTabColor(blueColor);

  Logger.log('Rack config sheet created: ' + sheetName);
  return newSheet;
}

/**
 * Pulls BOM from Arena for a rack and populates the rack config sheet
 * @param {Sheet} sheet - Rack config sheet to populate
 * @param {string} itemNumber - Arena item number to pull BOM from
 * @param {string} itemGuid - Optional Arena item GUID (if already known, skips item lookup)
 */
function pullBOMForRack(sheet, itemNumber, itemGuid) {
  try {
    var arenaClient = new ArenaAPIClient();

    Logger.log('pullBOMForRack: Starting BOM pull for item: ' + itemNumber);

    // If GUID is provided, use it directly; otherwise look up the item
    if (!itemGuid) {
      Logger.log('pullBOMForRack: GUID not provided, searching for item by number');
      var item = arenaClient.getItemByNumber(itemNumber);

      if (!item) {
        throw new Error('Item not found in Arena: ' + itemNumber);
      }

      itemGuid = item.guid || item.Guid;
    }

    Logger.log('pullBOMForRack: Using item GUID: ' + itemGuid);

    if (!itemGuid) {
      throw new Error('Item GUID is missing for item: ' + itemNumber);
    }

    Logger.log('pullBOMForRack: Fetching BOM for item: ' + itemNumber + ' (GUID: ' + itemGuid + ')');

    // Get BOM from Arena
    var bomData = arenaClient.makeRequest('/items/' + itemGuid + '/bom', { method: 'GET' });
    Logger.log('pullBOMForRack: BOM API response received');
    Logger.log('pullBOMForRack: Response structure: ' + JSON.stringify(Object.keys(bomData || {})));

    var bomLines = bomData.results || bomData.Results || [];

    Logger.log('pullBOMForRack: Retrieved ' + bomLines.length + ' BOM lines from Arena');

    if (bomLines.length === 0) {
      Logger.log('pullBOMForRack: No BOM lines found for item: ' + itemNumber);
      throw new Error('No BOM found for item ' + itemNumber + '. The item may not have any components.');
    }
  } catch (error) {
    Logger.log('pullBOMForRack: ERROR - ' + error.message);
    Logger.log('pullBOMForRack: ERROR Stack - ' + error.stack);
    throw error;
  }

  // Populate sheet with BOM data
  var rowData = [];
  var categoryColors = getCategoryColors();

  Logger.log('pullBOMForRack: Processing ' + bomLines.length + ' BOM lines');

  // Log sample BOM line structure
  if (bomLines.length > 0) {
    Logger.log('pullBOMForRack: Sample BOM line structure: ' + JSON.stringify(Object.keys(bomLines[0])));
    if (bomLines[0].item || bomLines[0].Item) {
      var sampleItem = bomLines[0].item || bomLines[0].Item;
      Logger.log('pullBOMForRack: Sample item structure: ' + JSON.stringify(Object.keys(sampleItem)));
    }
  }

  bomLines.forEach(function(line) {
    var bomItem = line.item || line.Item || {};
    var quantity = line.quantity || line.Quantity || 1;

    var bomItemNumber = bomItem.number || bomItem.Number || '';
    var bomItemName = bomItem.name || bomItem.Name || '';
    var bomItemDesc = bomItem.description || bomItem.Description || '';

    var categoryObj = bomItem.category || bomItem.Category || {};
    var categoryName = categoryObj.name || categoryObj.Name || '';

    var lifecycleObj = bomItem.lifecyclePhase || bomItem.LifecyclePhase || {};
    var lifecycleName = lifecycleObj.name || lifecycleObj.Name || '';

    // If description or category is missing, we may need to fetch full item details
    if (!bomItemDesc || !categoryName) {
      var bomItemGuid = bomItem.guid || bomItem.Guid;
      if (bomItemGuid) {
        try {
          Logger.log('pullBOMForRack: Fetching full details for item: ' + bomItemNumber + ' (GUID: ' + bomItemGuid + ')');
          var fullItem = arenaClient.getItem(bomItemGuid);

          if (fullItem) {
            bomItemDesc = bomItemDesc || fullItem.description || fullItem.Description || '';

            var fullCategoryObj = fullItem.category || fullItem.Category || {};
            categoryName = categoryName || fullCategoryObj.name || fullCategoryObj.Name || '';

            var fullLifecycleObj = fullItem.lifecyclePhase || fullItem.LifecyclePhase || {};
            lifecycleName = lifecycleName || fullLifecycleObj.name || fullLifecycleObj.Name || '';
          }
        } catch (itemError) {
          Logger.log('pullBOMForRack: Could not fetch full details for ' + bomItemNumber + ': ' + itemError.message);
        }
      }
    }

    rowData.push([
      bomItemNumber,
      bomItemName,
      bomItemDesc,
      categoryName,
      lifecycleName,
      quantity
    ]);
  });

  // Write BOM data to sheet (starting at row 3)
  if (rowData.length > 0) {
    sheet.getRange(3, 1, rowData.length, 6).setValues(rowData);

    // Apply category colors to each row
    for (var i = 0; i < rowData.length; i++) {
      var category = rowData[i][3];  // Category is column 4 (index 3)
      var color = getCategoryColor(category);

      if (color) {
        var rowNum = i + 3;  // Data starts at row 3
        var lastCol = sheet.getLastColumn();
        sheet.getRange(rowNum, 1, 1, lastCol).setBackground(color);
      }
    }

    Logger.log('Populated ' + rowData.length + ' BOM lines into rack config');

    // Update status to SYNCED now that BOM has been pulled from Arena
    Logger.log('pullBOMForRack: Updating rack status to SYNCED with GUID: ' + itemGuid);
    var eventDetails = {
      changesSummary: 'BOM pulled from Arena (' + rowData.length + ' items)',
      details: 'Initial BOM pull from Arena item ' + itemNumber
    };
    updateRackSheetStatus(sheet, RACK_STATUS.SYNCED, itemGuid, eventDetails);

    // Log BOM pull event
    addRackHistoryEvent(itemNumber, HISTORY_EVENT.BOM_PULL, {
      changesSummary: rowData.length + ' items pulled from Arena',
      details: 'Initial BOM synchronization from Arena',
      statusAfter: RACK_STATUS.SYNCED
    });
  }
}

/**
 * Navigates to Rack History tab and opens the filter sidebar
 * Menu item: Arena Data Center → View Rack History
 */
function viewRackHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var historySheet = getOrCreateRackHistoryTab();

    // Navigate to History tab
    ss.setActiveSheet(historySheet);

    // Clear cache to allow auto-open next time
    var cache = CacheService.getScriptCache();
    cache.remove('history_sidebar_shown');

    // Open sidebar
    showHistoryFilterSidebar();

  } catch (error) {
    Logger.log('Error in viewRackHistory: ' + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open Rack History: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Clears the sidebar cache (for troubleshooting)
 * Run this from Apps Script editor if auto-open stops working
 */
function clearSidebarCache() {
  var cache = CacheService.getScriptCache();
  cache.remove('history_sidebar_shown');
  Logger.log('Sidebar cache cleared - auto-open will work again');
  SpreadsheetApp.getUi().alert('Cache Cleared', 'Sidebar will auto-open next time you visit Rack History tab.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// TYPE SYSTEM CONFIGURATION FUNCTIONS (NEW)
// ============================================================================

/**
 * Shows the type system configuration UI
 * NEW: Allows users to configure entity types, categories, and terminology
 */
function showConfigureTypeSystem() {
  // TODO: Create ConfigureTypeSystem.html
  // For now, show a placeholder message
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'Type System Configuration',
    'Type system configuration UI coming soon!\n\n' +
    'For now, you can:\n' +
    '• Run the Setup Wizard to reconfigure\n' +
    '• Export/Import your configuration\n' +
    '• Use the migration functions',
    ui.ButtonSet.OK
  );
}

/**
 * Shows a dialog to reset the configuration system
 * NEW: Allows users to start fresh with default configuration
 */
function resetConfigurationDialog() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    'Reset Configuration',
    'This will reset your type system configuration to defaults.\n\n' +
    '⚠️ WARNING: This will erase your custom:\n' +
    '• Entity type names and keywords\n' +
    '• Category classifications\n' +
    '• Layout settings\n' +
    '• Hierarchy levels\n\n' +
    'Your sheets and data will NOT be affected.\n\n' +
    'Continue with reset?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    var result = resetTypeSystemConfiguration();

    if (result.success) {
      ui.alert(
        'Configuration Reset',
        'Type system configuration has been reset.\n\n' +
        'Please reload the spreadsheet to run the setup wizard again.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Reset Failed',
        'Could not reset configuration: ' + result.message,
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * Manually runs datacenter auto-migration
 * For users with existing datacenter spreadsheets
 */
function runManualMigration() {
  var ui = SpreadsheetApp.getUi();

  // Check if already initialized
  if (isSystemInitialized()) {
    ui.alert(
      'Already Initialized',
      'System is already configured.\n\n' +
      'Use Setup → Export Configuration to backup your current settings.',
      ui.ButtonSet.OK
    );
    return;
  }

  // Detect existing configuration
  var detection = detectExistingConfiguration();

  if (!detection.detected || detection.type !== 'datacenter') {
    ui.alert(
      'No Datacenter Configuration Detected',
      'Could not detect a datacenter configuration.\n\n' +
      detection.message + '\n\n' +
      'Use Setup → Quick Setup for default configuration instead.',
      ui.ButtonSet.OK
    );
    return;
  }

  // Confirm migration
  var response = ui.alert(
    'Run Auto-Migration',
    'Detected: ' + detection.message + '\n\n' +
    'This will migrate your datacenter configuration to the new type system.\n' +
    'All existing sheets and data will continue to work.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // Run migration
  var result = migrateFromV1();

  if (result.success) {
    ui.alert(
      'Migration Successful',
      result.message + '\n\n' +
      'Please reload the spreadsheet to see the updated menus.',
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      'Migration Failed',
      'Migration failed: ' + result.message + '\n\n' +
      'Please check the execution log for details.',
      ui.ButtonSet.OK
    );
  }
}
