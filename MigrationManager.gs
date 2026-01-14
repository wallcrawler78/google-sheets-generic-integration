/**
 * MigrationManager.gs
 * Handles detection and migration of existing configurations to the new type system.
 *
 * This module provides:
 * - Detection of existing datacenter configurations
 * - Migration from V1 (hardcoded) to V2 (configurable)
 * - Configuration export/import for backup and transfer
 */

// ============================================================================
// DETECTION FUNCTIONS
// ============================================================================

/**
 * Detects existing configuration type
 * Checks for datacenter-specific sheets to identify V1 installations
 * @return {Object} Detection result {detected: boolean, type: string, confidence: string}
 */
function detectExistingConfiguration() {
  // First check if already initialized with new system
  if (isSystemInitialized()) {
    return {
      detected: true,
      type: 'custom',
      confidence: 'high',
      message: 'System already configured with type system'
    };
  }

  // Check for datacenter-specific sheet names
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = spreadsheet.getSheets();
  var sheetNames = allSheets.map(function(sheet) {
    return sheet.getName();
  });

  // Look for characteristic datacenter sheets
  var datacenterIndicators = {
    'Legend-NET': false,
    'Overhead': false,
    'Full Rack Item A': false,
    'Full Rack Item B': false,
    'Full Rack Item C': false
  };

  for (var i = 0; i < sheetNames.length; i++) {
    var sheetName = sheetNames[i];
    if (datacenterIndicators.hasOwnProperty(sheetName)) {
      datacenterIndicators[sheetName] = true;
    }
  }

  // Count how many datacenter indicators were found
  var indicatorCount = 0;
  for (var key in datacenterIndicators) {
    if (datacenterIndicators[key]) {
      indicatorCount++;
    }
  }

  if (indicatorCount >= 3) {
    // High confidence datacenter detection
    return {
      detected: true,
      type: 'datacenter',
      confidence: 'high',
      message: 'Detected datacenter configuration with ' + indicatorCount + ' characteristic sheets'
    };
  } else if (indicatorCount >= 1) {
    // Medium confidence
    return {
      detected: true,
      type: 'datacenter',
      confidence: 'medium',
      message: 'Possible datacenter configuration with ' + indicatorCount + ' characteristic sheets'
    };
  }

  // No existing configuration detected
  return {
    detected: false,
    type: null,
    confidence: 'none',
    message: 'No existing configuration detected'
  };
}

// ============================================================================
// MIGRATION FUNCTIONS
// ============================================================================

/**
 * Migrates from V1 datacenter configuration to new type system
 * Converts hardcoded datacenter terminology to configurable system
 * @return {Object} Migration result {success: boolean, message: string}
 */
function migrateFromV1() {
  try {
    Logger.log('Starting V1 datacenter migration...');

    // Create datacenter configuration
    var config = {
      primaryEntity: getDatacenterPrimaryEntityType(),
      typeDefinitions: getDatacenterTypeDefinitions(),
      categoryClassifications: getDatacenterCategoryClassifications(),
      layoutConfig: getDatacenterLayoutConfig(),
      hierarchyLevels: getDatacenterHierarchyLevels()
    };

    // Validate configuration before saving
    var validation = validateTypeSystemConfiguration(config);
    if (!validation.valid) {
      Logger.log('Migration validation failed: ' + validation.errors.join(', '));
      return {
        success: false,
        message: 'Migration validation failed: ' + validation.errors.join(', ')
      };
    }

    // Save configuration
    var saveResult = saveTypeSystemConfiguration(config);
    if (!saveResult.success) {
      return saveResult;
    }

    // Mark as migrated from V1
    PropertiesService.getScriptProperties().setProperty(CONFIG_KEYS.MIGRATED_FROM_V1, 'true');

    Logger.log('V1 migration completed successfully');

    return {
      success: true,
      message: 'Successfully migrated datacenter configuration to new type system'
    };

  } catch (e) {
    Logger.log('Error during V1 migration: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);
    return {
      success: false,
      message: 'Migration error: ' + e.message
    };
  }
}

/**
 * Checks if this system was migrated from V1
 * @return {boolean} True if migrated from V1
 */
function wasMigratedFromV1() {
  var migrated = PropertiesService.getScriptProperties().getProperty(CONFIG_KEYS.MIGRATED_FROM_V1);
  return migrated === 'true';
}

/**
 * Performs silent auto-migration if needed
 * Called from onOpen() to transparently migrate existing spreadsheets
 * @return {Object} Migration result {performed: boolean, success: boolean, message: string}
 */
function autoMigrateIfNeeded() {
  // Check if already initialized
  if (isSystemInitialized()) {
    return {
      performed: false,
      success: true,
      message: 'System already initialized, no migration needed'
    };
  }

  // Detect existing configuration
  var detection = detectExistingConfiguration();

  if (detection.detected && detection.type === 'datacenter') {
    // Perform migration
    Logger.log('Auto-migration triggered: ' + detection.message);
    var result = migrateFromV1();

    return {
      performed: true,
      success: result.success,
      message: result.message,
      showNotification: result.success // Show notification to user if successful
    };
  }

  // No migration needed
  return {
    performed: false,
    success: true,
    message: 'No migration needed'
  };
}

// ============================================================================
// EXPORT/IMPORT FUNCTIONS
// ============================================================================

/**
 * Exports current configuration as JSON string
 * Useful for backup and transferring configurations between spreadsheets
 * @return {Object} Export result {success: boolean, json: string, message: string}
 */
function exportConfiguration() {
  try {
    if (!isSystemInitialized()) {
      return {
        success: false,
        json: null,
        message: 'No configuration to export - system not initialized'
      };
    }

    var config = {
      primaryEntity: getPrimaryEntityType(),
      typeDefinitions: getTypeDefinitions(),
      categoryClassifications: getCategoryClassifications(),
      layoutConfig: getLayoutConfig(),
      hierarchyLevels: getHierarchyLevels(),
      exportedAt: new Date().toISOString(),
      version: '2.0'
    };

    var json = JSON.stringify(config, null, 2);

    Logger.log('Configuration exported successfully');

    return {
      success: true,
      json: json,
      message: 'Configuration exported successfully'
    };

  } catch (e) {
    Logger.log('Error exporting configuration: ' + e.message);
    return {
      success: false,
      json: null,
      message: 'Export error: ' + e.message
    };
  }
}

/**
 * Imports configuration from JSON string
 * Validates and saves the imported configuration
 * @param {string} json - JSON string containing configuration
 * @return {Object} Import result {success: boolean, message: string}
 */
function importConfiguration(json) {
  try {
    // Parse JSON
    var config = JSON.parse(json);

    // Validate configuration
    var validation = validateTypeSystemConfiguration(config);
    if (!validation.valid) {
      return {
        success: false,
        message: 'Invalid configuration: ' + validation.errors.join(', ')
      };
    }

    // Save configuration
    var saveResult = saveTypeSystemConfiguration(config);
    if (!saveResult.success) {
      return saveResult;
    }

    Logger.log('Configuration imported successfully');

    return {
      success: true,
      message: 'Configuration imported successfully'
    };

  } catch (e) {
    Logger.log('Error importing configuration: ' + e.message);
    return {
      success: false,
      message: 'Import error: ' + e.message
    };
  }
}

/**
 * Creates a downloadable JSON file of the current configuration
 * Opens a dialog showing the JSON that user can copy
 */
function showExportDialog() {
  var exportResult = exportConfiguration();

  if (!exportResult.success) {
    SpreadsheetApp.getUi().alert('Export Failed', exportResult.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Create HTML output to display JSON
  // SECURITY FIX: Use JavaScript to set textarea value to prevent XSS
  var html = HtmlService.createHtmlOutput(
    '<h3>Configuration Export</h3>' +
    '<p>Copy the JSON below to back up your configuration:</p>' +
    '<textarea id="configJson" style="width:100%; height:400px; font-family:monospace; font-size:12px;"></textarea>' +
    '<br><br>' +
    '<button onclick="selectAll()">Select All</button>' +
    '<button onclick="google.script.host.close()">Close</button>' +
    '<script>' +
    '(function() {' +
    '  var jsonData = ' + JSON.stringify(exportResult.json) + ';' +
    '  document.getElementById("configJson").value = jsonData;' +
    '})();' +
    'function selectAll() {' +
    '  document.getElementById("configJson").select();' +
    '  document.execCommand("copy");' +
    '  alert("Copied to clipboard!");' +
    '}' +
    '</script>'
  )
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Export Configuration');
}

/**
 * Shows import dialog for pasting JSON configuration
 */
function showImportDialog() {
  var html = HtmlService.createHtmlOutput(
    '<h3>Configuration Import</h3>' +
    '<p>Paste your configuration JSON below:</p>' +
    '<textarea id="configJson" style="width:100%; height:400px; font-family:monospace; font-size:12px;"></textarea>' +
    '<br><br>' +
    '<button onclick="importConfig()">Import</button>' +
    '<button onclick="google.script.host.close()">Cancel</button>' +
    '<div id="result" style="margin-top:10px; padding:10px; display:none;"></div>' +
    '<script>' +
    'function importConfig() {' +
    '  var json = document.getElementById("configJson").value;' +
    '  if (!json.trim()) {' +
    '    alert("Please paste a configuration JSON");' +
    '    return;' +
    '  }' +
    '  google.script.run' +
    '    .withSuccessHandler(function(result) {' +
    '      var resultDiv = document.getElementById("result");' +
    '      resultDiv.style.display = "block";' +
    '      if (result.success) {' +
    '        resultDiv.style.backgroundColor = "#d4edda";' +
    '        resultDiv.style.color = "#155724";' +
    '        resultDiv.innerHTML = "Success: " + result.message;' +
    '        setTimeout(function() { google.script.host.close(); }, 2000);' +
    '      } else {' +
    '        resultDiv.style.backgroundColor = "#f8d7da";' +
    '        resultDiv.style.color = "#721c24";' +
    '        resultDiv.innerHTML = "Error: " + result.message;' +
    '      }' +
    '    })' +
    '    .withFailureHandler(function(error) {' +
    '      var resultDiv = document.getElementById("result");' +
    '      resultDiv.style.display = "block";' +
    '      resultDiv.style.backgroundColor = "#f8d7da";' +
    '      resultDiv.style.color = "#721c24";' +
    '      resultDiv.innerHTML = "Error: " + error.message;' +
    '    })' +
    '    .importConfiguration(json);' +
    '}' +
    '</script>'
  )
    .setWidth(600)
    .setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, 'Import Configuration');
}

// ============================================================================
// MIGRATION NOTIFICATION
// ============================================================================

/**
 * Shows a one-time notification to users after auto-migration
 * Informs them about the new configuration system
 */
function showMigrationNotification() {
  var ui = SpreadsheetApp.getUi();

  var message =
    'Your datacenter configuration has been migrated to the new type system.\n\n' +
    'All your existing sheets and data continue to work exactly as before.\n\n' +
    'NEW: You can now customize terminology and classifications:\n' +
    '• Go to Setup → Configure Type System\n' +
    '• Rename "Rack" to any term you prefer\n' +
    '• Add or modify type classifications\n' +
    '• Change category keywords and colors\n\n' +
    'This notification will only appear once.';

  ui.alert(
    'Configuration System Updated',
    message,
    ui.ButtonSet.OK
  );

  // Mark notification as shown
  PropertiesService.getUserProperties().setProperty('MIGRATION_NOTIFICATION_SHOWN', 'true');
}

/**
 * Checks if migration notification should be shown
 * @return {boolean} True if notification should be shown
 */
function shouldShowMigrationNotification() {
  // Only show if:
  // 1. System was migrated from V1
  // 2. Notification hasn't been shown yet
  var wasMigrated = wasMigratedFromV1();
  var notificationShown = PropertiesService.getUserProperties().getProperty('MIGRATION_NOTIFICATION_SHOWN') === 'true';

  return wasMigrated && !notificationShown;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Gets migration status information
 * Useful for debugging and support
 * @return {Object} Migration status
 */
function getMigrationStatus() {
  return {
    systemInitialized: isSystemInitialized(),
    migratedFromV1: wasMigratedFromV1(),
    detection: detectExistingConfiguration(),
    notificationShown: PropertiesService.getUserProperties().getProperty('MIGRATION_NOTIFICATION_SHOWN') === 'true'
  };
}

/**
 * Logs migration status to console
 * For debugging purposes
 */
function logMigrationStatus() {
  var status = getMigrationStatus();
  Logger.log('=== Migration Status ===');
  Logger.log('System Initialized: ' + status.systemInitialized);
  Logger.log('Migrated from V1: ' + status.migratedFromV1);
  Logger.log('Detection Result: ' + JSON.stringify(status.detection));
  Logger.log('Notification Shown: ' + status.notificationShown);
}
