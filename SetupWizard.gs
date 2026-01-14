/**
 * SetupWizard.gs
 * Server-side functions for the setup wizard.
 *
 * This module provides:
 * - Setup wizard display functions
 * - Configuration initialization
 * - First-run detection
 * - Initial sheet structure creation
 */

// ============================================================================
// WIZARD DISPLAY FUNCTIONS
// ============================================================================

/**
 * Shows the setup wizard modal dialog
 * Called automatically on first run or manually from menu
 */
function showSetupWizard() {
  var html = HtmlService.createHtmlOutputFromFile('SetupWizard')
    .setWidth(900)
    .setHeight(650)
    .setTitle('Initial Setup Wizard');

  SpreadsheetApp.getUi().showModalDialog(html, 'Setup Wizard');
}

/**
 * Shows a simplified quick setup dialog for users who want defaults
 */
function showQuickSetup() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
    'Quick Setup',
    'Would you like to use default neutral terminology?\n\n' +
    '• Entity Type: "Configuration"\n' +
    '• 3 standard types\n' +
    '• 4x10 grid layout\n\n' +
    'You can customize everything later in Setup → Configure Type System.',
    ui.ButtonSet.YES_NO
  );

  if (result === ui.Button.YES) {
    // Use default neutral configuration
    var config = {
      primaryEntity: getDefaultPrimaryEntityType(),
      typeDefinitions: getDefaultTypeDefinitions(),
      categoryClassifications: [],
      layoutConfig: getDefaultLayoutConfig(),
      hierarchyLevels: getDefaultHierarchyLevels()
    };

    var saveResult = saveSetupWizardConfig(config);

    if (saveResult.success) {
      ui.alert('Setup Complete', 'Default configuration applied successfully!', ui.ButtonSet.OK);
      // Reload to show new menu
      SpreadsheetApp.flush();
    } else {
      ui.alert('Setup Failed', saveResult.message, ui.ButtonSet.OK);
    }
  } else {
    // Show full wizard
    showSetupWizard();
  }
}

// ============================================================================
// DATA LOADING FUNCTIONS
// ============================================================================

/**
 * Loads data for setup wizard
 * Provides defaults and migration detection
 * @return {Object} Wizard data
 */
function loadSetupWizardData() {
  var existingConfig = detectExistingConfiguration();

  return {
    // Migration detection
    hasExistingConfig: existingConfig.detected,
    existingConfigType: existingConfig.type,
    existingConfigMessage: existingConfig.message,

    // Default configurations
    defaultPrimaryEntity: getDefaultPrimaryEntityType(),
    defaultTypeDefinitions: getDefaultTypeDefinitions(),
    defaultCategoryClassifications: [],
    defaultLayoutConfig: getDefaultLayoutConfig(),
    defaultHierarchyLevels: getDefaultHierarchyLevels(),

    // Datacenter configurations (for migration option)
    datacenterPrimaryEntity: getDatacenterPrimaryEntityType(),
    datacenterTypeDefinitions: getDatacenterTypeDefinitions(),
    datacenterCategoryClassifications: getDatacenterCategoryClassifications(),
    datacenterLayoutConfig: getDatacenterLayoutConfig(),
    datacenterHierarchyLevels: getDatacenterHierarchyLevels()
  };
}

// ============================================================================
// CONFIGURATION SAVE FUNCTIONS
// ============================================================================

/**
 * Saves configuration from setup wizard
 * Validates, saves, and creates initial structure
 * @param {Object} config - Complete configuration from wizard
 * @return {Object} Save result {success: boolean, message: string, errors: Array}
 */
function saveSetupWizardConfig(config) {
  try {
    Logger.log('Saving setup wizard configuration...');

    // Validate configuration
    var validation = validateTypeSystemConfiguration(config);
    if (!validation.valid) {
      Logger.log('Validation failed: ' + validation.errors.join(', '));
      return {
        success: false,
        message: 'Configuration validation failed',
        errors: validation.errors
      };
    }

    // Save configuration
    var saveResult = saveTypeSystemConfiguration(config);
    if (!saveResult.success) {
      return {
        success: false,
        message: saveResult.message,
        errors: [saveResult.message]
      };
    }

    // Create initial sheet structure
    try {
      createInitialSheetStructure(config);
    } catch (e) {
      Logger.log('Warning: Error creating initial sheets (non-fatal): ' + e.message);
      // Non-fatal - configuration is saved, sheets can be created later
    }

    Logger.log('Setup wizard configuration saved successfully');

    return {
      success: true,
      message: 'Configuration saved successfully! Reloading...',
      errors: []
    };

  } catch (e) {
    Logger.log('Error saving setup wizard config: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);
    return {
      success: false,
      message: 'Error saving configuration: ' + e.message,
      errors: [e.message]
    };
  }
}

// ============================================================================
// INITIAL STRUCTURE CREATION
// ============================================================================

/**
 * Creates initial sheet structure based on configuration
 * Creates overview and legend sheets
 * @param {Object} config - Configuration object
 */
function createInitialSheetStructure(config) {
  Logger.log('Creating initial sheet structure...');

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Determine overview sheet name from hierarchy
  var overviewName = 'Overview';
  if (config.hierarchyLevels && config.hierarchyLevels.length > 0) {
    overviewName = config.hierarchyLevels[0].name + ' Overview';
  }

  // Create or get overview sheet
  var overviewSheet = spreadsheet.getSheetByName(overviewName);
  if (!overviewSheet) {
    overviewSheet = spreadsheet.insertSheet(overviewName);
    Logger.log('Created overview sheet: ' + overviewName);
  }

  // Create or get legend sheet
  var legendSheet = spreadsheet.getSheetByName('Legend');
  if (!legendSheet) {
    legendSheet = spreadsheet.insertSheet('Legend');
    Logger.log('Created legend sheet');
  }

  // Set up overview sheet with basic structure if empty
  if (overviewSheet.getLastRow() === 0) {
    setupOverviewSheet(overviewSheet, config);
  }

  // Set up legend sheet if empty
  if (legendSheet.getLastRow() === 0) {
    setupLegendSheet(legendSheet, config);
  }

  Logger.log('Initial sheet structure created');
}

/**
 * Sets up the overview sheet with headers and structure
 * @param {Sheet} sheet - The overview sheet
 * @param {Object} config - Configuration object
 */
function setupOverviewSheet(sheet, config) {
  // Title row
  var primaryEntity = config.primaryEntity || getDefaultPrimaryEntityType();
  var hierarchyLevel0 = (config.hierarchyLevels && config.hierarchyLevels.length > 0) ?
    config.hierarchyLevels[0].name : 'Assembly';

  sheet.getRange(1, 1).setValue(hierarchyLevel0 + ' Overview');
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');

  // Instructions
  sheet.getRange(3, 1).setValue('Use Layout → Create New Overview Layout to set up your grid');
  sheet.getRange(3, 1).setFontStyle('italic');

  Logger.log('Overview sheet structure initialized');
}

/**
 * Sets up the legend sheet with category information
 * @param {Sheet} sheet - The legend sheet
 * @param {Object} config - Configuration object
 */
function setupLegendSheet(sheet, config) {
  // Title
  sheet.getRange(1, 1).setValue('Type & Category Legend');
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');

  // Type Definitions section
  sheet.getRange(3, 1).setValue('Type Classifications:');
  sheet.getRange(3, 1).setFontWeight('bold');

  var row = 4;
  var typeDefinitions = config.typeDefinitions || [];
  for (var i = 0; i < typeDefinitions.length; i++) {
    var typeDef = typeDefinitions[i];
    if (typeDef.enabled) {
      sheet.getRange(row, 1).setValue(typeDef.name);
      sheet.getRange(row, 2).setValue(typeDef.keywords.join(', '));
      sheet.getRange(row, 1).setBackground(typeDef.color);
      row++;
    }
  }

  // Category Classifications section
  if (config.categoryClassifications && config.categoryClassifications.length > 0) {
    row += 2;
    sheet.getRange(row, 1).setValue('Category Classifications:');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;

    for (var j = 0; j < config.categoryClassifications.length; j++) {
      var catClass = config.categoryClassifications[j];
      if (catClass.enabled) {
        sheet.getRange(row, 1).setValue(catClass.name);
        sheet.getRange(row, 2).setValue(catClass.keywords.join(', '));
        sheet.getRange(row, 1).setBackground(catClass.color);
        row++;
      }
    }
  }

  // Auto-resize columns
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);

  Logger.log('Legend sheet structure initialized');
}

// ============================================================================
// FIRST RUN DETECTION
// ============================================================================

/**
 * Checks if this is the first run and handles initialization
 * Called from onOpen() in Code.gs
 * @return {Object} First run status {isFirstRun: boolean, action: string}
 */
function checkFirstRun() {
  // Check if system is initialized
  if (isSystemInitialized()) {
    return {
      isFirstRun: false,
      action: 'none',
      message: 'System already initialized'
    };
  }

  // Check for existing configuration to migrate
  var detection = detectExistingConfiguration();

  if (detection.detected && detection.type === 'datacenter') {
    // Auto-migrate datacenter configuration
    return {
      isFirstRun: false,
      action: 'auto-migrate',
      message: 'Datacenter configuration detected, will auto-migrate',
      detection: detection
    };
  }

  // This is a true first run - show setup wizard
  return {
    isFirstRun: true,
    action: 'show-wizard',
    message: 'First run detected, show setup wizard',
    detection: detection
  };
}

// ============================================================================
// EXAMPLE CONFIGURATIONS
// ============================================================================

/**
 * Gets example configurations for different use cases
 * Helps users understand what they can configure
 * @return {Object} Examples for different industries/use cases
 */
function getExampleConfigurations() {
  return {
    manufacturing: {
      name: 'Manufacturing / Assembly',
      primaryEntity: {
        singular: 'Assembly',
        plural: 'Assemblies',
        verb: 'Assemble'
      },
      typeDefinitions: [
        {
          id: 'TYPE_1',
          name: 'Standard Assembly',
          keywords: ['standard', 'basic'],
          color: '#00FFFF',
          enabled: true
        },
        {
          id: 'TYPE_2',
          name: 'Custom Assembly',
          keywords: ['custom', 'special'],
          color: '#FFA500',
          enabled: true
        }
      ]
    },
    it: {
      name: 'IT / Servers',
      primaryEntity: {
        singular: 'Server',
        plural: 'Servers',
        verb: 'Server'
      },
      typeDefinitions: [
        {
          id: 'TYPE_1',
          name: 'Web Server',
          keywords: ['web', 'apache', 'nginx'],
          color: '#00FFFF',
          enabled: true
        },
        {
          id: 'TYPE_2',
          name: 'Database Server',
          keywords: ['database', 'sql', 'mysql'],
          color: '#FFA500',
          enabled: true
        },
        {
          id: 'TYPE_3',
          name: 'Application Server',
          keywords: ['application', 'app'],
          color: '#00FF00',
          enabled: true
        }
      ]
    },
    retail: {
      name: 'Retail / Products',
      primaryEntity: {
        singular: 'Product',
        plural: 'Products',
        verb: 'Catalog'
      },
      typeDefinitions: [
        {
          id: 'TYPE_1',
          name: 'Electronics',
          keywords: ['electronic', 'tech', 'digital'],
          color: '#00FFFF',
          enabled: true
        },
        {
          id: 'TYPE_2',
          name: 'Clothing',
          keywords: ['clothing', 'apparel', 'fashion'],
          color: '#FFA500',
          enabled: true
        },
        {
          id: 'TYPE_3',
          name: 'Home Goods',
          keywords: ['home', 'furniture', 'decor'],
          color: '#00FF00',
          enabled: true
        }
      ]
    }
  };
}

// ============================================================================
// VALIDATION HELPERS
// ============================================================================

/**
 * Validates a single type definition
 * @param {Object} typeDef - Type definition to validate
 * @return {Object} Validation result {valid: boolean, errors: Array}
 */
function validateTypeDefinition(typeDef) {
  var errors = [];

  if (!typeDef.name || typeDef.name.trim() === '') {
    errors.push('Type name is required');
  }

  if (!typeDef.keywords || typeDef.keywords.length === 0) {
    errors.push('At least one keyword is required');
  }

  if (!typeDef.color || !typeDef.color.match(/^#[0-9A-Fa-f]{6}$/)) {
    errors.push('Valid hex color is required (e.g., #00FFFF)');
  }

  return {
    valid: errors.length === 0,
    errors: errors
  };
}

/**
 * Validates a single category classification
 * @param {Object} catClass - Category classification to validate
 * @return {Object} Validation result {valid: boolean, errors: Array}
 */
function validateCategoryClassification(catClass) {
  var errors = [];

  if (!catClass.name || catClass.name.trim() === '') {
    errors.push('Category name is required');
  }

  if (!catClass.keywords || catClass.keywords.length === 0) {
    errors.push('At least one keyword is required');
  }

  if (!catClass.color || !catClass.color.match(/^#[0-9A-Fa-f]{6}$/)) {
    errors.push('Valid hex color is required (e.g., #FFFF00)');
  }

  return {
    valid: errors.length === 0,
    errors: errors
  };
}
