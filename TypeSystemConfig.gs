/**
 * TypeSystemConfig.gs
 * Configuration management for the generic type classification system.
 * Replaces hardcoded datacenter-specific terminology with user-configurable types.
 *
 * This module provides:
 * - Configuration storage/retrieval via PropertiesService
 * - Default configurations (neutral and datacenter)
 * - Validation functions
 * - Terminology helper functions
 */

// ============================================================================
// CONFIGURATION KEYS
// ============================================================================

var CONFIG_KEYS = {
  // Core system configuration
  SYSTEM_INITIALIZED: 'SYSTEM_INITIALIZED',
  PRIMARY_ENTITY_TYPE: 'PRIMARY_ENTITY_TYPE',
  TYPE_DEFINITIONS: 'TYPE_DEFINITIONS',
  CATEGORY_CLASSIFICATIONS: 'CATEGORY_CLASSIFICATIONS',
  LAYOUT_CONFIG: 'LAYOUT_CONFIG',
  HIERARCHY_LEVELS: 'HIERARCHY_LEVELS',

  // Migration flag for existing users
  MIGRATED_FROM_V1: 'MIGRATED_FROM_V1'
};

// ============================================================================
// GETTER FUNCTIONS
// ============================================================================

/**
 * Gets the primary entity type configuration
 * @return {Object} Primary entity configuration {singular, plural, verb}
 */
function getPrimaryEntityType() {
  var json = PropertiesService.getScriptProperties().getProperty(CONFIG_KEYS.PRIMARY_ENTITY_TYPE);
  if (!json) {
    // Return neutral defaults if not configured
    return getDefaultPrimaryEntityType();
  }
  try {
    return JSON.parse(json);
  } catch (e) {
    Logger.log('Error parsing PRIMARY_ENTITY_TYPE: ' + e.message);
    return getDefaultPrimaryEntityType();
  }
}

/**
 * Gets user-configured type definitions
 * @return {Array<Object>} Type definitions array
 */
function getTypeDefinitions() {
  var json = PropertiesService.getScriptProperties().getProperty(CONFIG_KEYS.TYPE_DEFINITIONS);
  if (!json) {
    // Return neutral defaults if not configured
    return getDefaultTypeDefinitions();
  }
  try {
    return JSON.parse(json);
  } catch (e) {
    Logger.log('Error parsing TYPE_DEFINITIONS: ' + e.message);
    return getDefaultTypeDefinitions();
  }
}

/**
 * Gets user-configured category classifications
 * @return {Array<Object>} Category classifications array
 */
function getCategoryClassifications() {
  var json = PropertiesService.getScriptProperties().getProperty(CONFIG_KEYS.CATEGORY_CLASSIFICATIONS);
  if (!json) {
    // Return empty array if not configured - categories are optional
    return [];
  }
  try {
    return JSON.parse(json);
  } catch (e) {
    Logger.log('Error parsing CATEGORY_CLASSIFICATIONS: ' + e.message);
    return [];
  }
}

/**
 * Gets layout configuration
 * @return {Object} Layout configuration
 */
function getLayoutConfig() {
  var json = PropertiesService.getScriptProperties().getProperty(CONFIG_KEYS.LAYOUT_CONFIG);
  if (!json) {
    return getDefaultLayoutConfig();
  }
  try {
    return JSON.parse(json);
  } catch (e) {
    Logger.log('Error parsing LAYOUT_CONFIG: ' + e.message);
    return getDefaultLayoutConfig();
  }
}

/**
 * Gets hierarchy levels configuration
 * @return {Array<Object>} Hierarchy levels array
 */
function getHierarchyLevels() {
  var json = PropertiesService.getScriptProperties().getProperty(CONFIG_KEYS.HIERARCHY_LEVELS);
  if (!json) {
    return getDefaultHierarchyLevels();
  }
  try {
    return JSON.parse(json);
  } catch (e) {
    Logger.log('Error parsing HIERARCHY_LEVELS: ' + e.message);
    return getDefaultHierarchyLevels();
  }
}

/**
 * Gets hierarchy level name by level number
 * @param {number} level - The level number (0, 1, 2, etc.)
 * @return {string} The level name (e.g., "Assembly", "POD")
 */
function getHierarchyLevelName(level) {
  var hierarchyLevels = getHierarchyLevels();
  for (var i = 0; i < hierarchyLevels.length; i++) {
    if (hierarchyLevels[i].level === level) {
      return hierarchyLevels[i].name;
    }
  }
  return 'Level ' + level; // Fallback
}

// ============================================================================
// SETTER FUNCTIONS
// ============================================================================

/**
 * Saves complete type system configuration
 * @param {Object} config - Complete configuration object
 * @return {Object} Save result {success: boolean, message: string}
 */
function saveTypeSystemConfiguration(config) {
  try {
    var scriptProps = PropertiesService.getScriptProperties();

    // Save each configuration section
    if (config.primaryEntity) {
      scriptProps.setProperty(CONFIG_KEYS.PRIMARY_ENTITY_TYPE, JSON.stringify(config.primaryEntity));
    }

    if (config.typeDefinitions) {
      scriptProps.setProperty(CONFIG_KEYS.TYPE_DEFINITIONS, JSON.stringify(config.typeDefinitions));
    }

    if (config.categoryClassifications) {
      scriptProps.setProperty(CONFIG_KEYS.CATEGORY_CLASSIFICATIONS, JSON.stringify(config.categoryClassifications));
    }

    if (config.layoutConfig) {
      scriptProps.setProperty(CONFIG_KEYS.LAYOUT_CONFIG, JSON.stringify(config.layoutConfig));
    }

    if (config.hierarchyLevels) {
      scriptProps.setProperty(CONFIG_KEYS.HIERARCHY_LEVELS, JSON.stringify(config.hierarchyLevels));
    }

    // Mark system as initialized
    scriptProps.setProperty(CONFIG_KEYS.SYSTEM_INITIALIZED, 'true');

    Logger.log('Type system configuration saved successfully');

    return {
      success: true,
      message: 'Configuration saved successfully'
    };

  } catch (e) {
    Logger.log('Error saving type system configuration: ' + e.message);
    return {
      success: false,
      message: 'Error saving configuration: ' + e.message
    };
  }
}

// ============================================================================
// DEFAULT CONFIGURATIONS - NEUTRAL (for new users)
// ============================================================================

/**
 * Gets default primary entity type (neutral terminology)
 * @return {Object} Default primary entity configuration
 */
function getDefaultPrimaryEntityType() {
  return {
    singular: 'Configuration',
    plural: 'Configurations',
    verb: 'Configure'
  };
}

/**
 * Gets default type definitions (neutral terminology)
 * @return {Array<Object>} Default type definitions
 */
function getDefaultTypeDefinitions() {
  return [
    {
      id: 'TYPE_1',
      name: 'Standard Configuration',
      displayName: 'Type A',
      keywords: ['standard', 'type a', 'config a'],
      color: '#00FFFF',
      enabled: true
    },
    {
      id: 'TYPE_2',
      name: 'Advanced Configuration',
      displayName: 'Type B',
      keywords: ['advanced', 'type b', 'config b'],
      color: '#FFA500',
      enabled: true
    },
    {
      id: 'TYPE_3',
      name: 'Premium Configuration',
      displayName: 'Type C',
      keywords: ['premium', 'type c', 'config c'],
      color: '#00FF00',
      enabled: true
    }
  ];
}

/**
 * Gets default layout configuration
 * @return {Object} Default layout configuration
 */
function getDefaultLayoutConfig() {
  return {
    gridType: 'OVERHEAD',
    rows: 4,
    positionsPerRow: 10,
    labelPrefix: 'Position',
    cellHeight: 25,
    cellWidth: 120
  };
}

/**
 * Gets default hierarchy levels (neutral terminology)
 * @return {Array<Object>} Default hierarchy levels
 */
function getDefaultHierarchyLevels() {
  return [
    {
      level: 0,
      name: 'Assembly',
      attribute: null
    },
    {
      level: 1,
      name: 'SubAssembly',
      attribute: null
    },
    {
      level: 2,
      name: 'Configuration',
      attribute: null
    }
  ];
}

// ============================================================================
// DEFAULT CONFIGURATIONS - DATACENTER (for migration)
// ============================================================================

/**
 * Gets datacenter-specific primary entity type (for migration)
 * @return {Object} Datacenter primary entity configuration
 */
function getDatacenterPrimaryEntityType() {
  return {
    singular: 'Rack',
    plural: 'Racks',
    verb: 'Rack'
  };
}

/**
 * Gets datacenter-specific type definitions (for migration)
 * @return {Array<Object>} Datacenter type definitions
 */
function getDatacenterTypeDefinitions() {
  return [
    {
      id: 'TYPE_A',
      name: 'Full Rack Item A',
      displayName: 'Rack A',
      keywords: ['rack a', 'type a', 'config a'],
      color: '#00FFFF',
      enabled: true
    },
    {
      id: 'TYPE_B',
      name: 'Full Rack Item B',
      displayName: 'Rack B',
      keywords: ['rack b', 'type b', 'config b'],
      color: '#FFA500',
      enabled: true
    },
    {
      id: 'TYPE_C',
      name: 'Full Rack Item C',
      displayName: 'Rack C',
      keywords: ['rack c', 'type c', 'config c'],
      color: '#00FF00',
      enabled: true
    },
    {
      id: 'TYPE_D',
      name: 'Full Rack Item D',
      displayName: 'Rack D',
      keywords: ['rack d', 'type d', 'config d'],
      color: '#90EE90',
      enabled: true
    },
    {
      id: 'TYPE_E',
      name: 'Full Rack Item E',
      displayName: 'Rack E',
      keywords: ['rack e', 'type e', 'config e'],
      color: '#90EE90',
      enabled: true
    },
    {
      id: 'TYPE_F',
      name: 'Full Rack F',
      displayName: 'Rack F',
      keywords: ['rack f', 'type f', 'config f'],
      color: '#0000FF',
      enabled: true
    },
    {
      id: 'TYPE_G',
      name: 'Full Rack Item G',
      displayName: 'Rack G',
      keywords: ['rack g', 'type g', 'config g'],
      color: '#000000',
      enabled: true
    }
  ];
}

/**
 * Gets datacenter-specific category classifications (for migration)
 * @return {Array<Object>} Datacenter category classifications
 */
function getDatacenterCategoryClassifications() {
  return [
    {
      id: 'CAT_ETH',
      name: 'Data center ETH',
      keywords: ['ethernet', 'eth', 'network switch', 'nic'],
      color: '#FFFF00',
      enabled: true
    },
    {
      id: 'CAT_SPINE',
      name: 'Data Center SPINE RACK',
      keywords: ['spine', 'backbone', 'core switch'],
      color: '#FF00FF',
      enabled: true
    },
    {
      id: 'CAT_GRID_POD',
      name: 'DATA CENTER GRID-POD',
      keywords: ['grid', 'pod', 'power distribution'],
      color: '#00FF00',
      enabled: true
    }
  ];
}

/**
 * Gets datacenter-specific layout configuration (for migration)
 * @return {Object} Datacenter layout configuration
 */
function getDatacenterLayoutConfig() {
  return {
    gridType: 'OVERHEAD',
    rows: 4,
    positionsPerRow: 10,
    labelPrefix: 'Pos',
    cellHeight: 25,
    cellWidth: 120
  };
}

/**
 * Gets datacenter-specific hierarchy levels (for migration)
 * @return {Array<Object>} Datacenter hierarchy levels
 */
function getDatacenterHierarchyLevels() {
  return [
    {
      level: 0,
      name: 'POD',
      attribute: null
    },
    {
      level: 1,
      name: 'Row',
      attribute: 'Row Location'
    },
    {
      level: 2,
      name: 'Rack',
      attribute: null
    }
  ];
}

// ============================================================================
// VALIDATION FUNCTIONS
// ============================================================================

/**
 * Validates type system configuration
 * @param {Object} config - Configuration to validate
 * @return {Object} Validation result {valid: boolean, errors: Array<string>}
 */
function validateTypeSystemConfiguration(config) {
  var errors = [];

  // Validate primary entity
  if (!config.primaryEntity) {
    errors.push('Primary entity configuration is required');
  } else {
    if (!config.primaryEntity.singular || config.primaryEntity.singular.trim() === '') {
      errors.push('Primary entity singular name is required');
    }
    if (!config.primaryEntity.plural || config.primaryEntity.plural.trim() === '') {
      errors.push('Primary entity plural name is required');
    }
  }

  // Validate type definitions
  if (!config.typeDefinitions || config.typeDefinitions.length === 0) {
    errors.push('At least one type definition is required');
  } else {
    for (var i = 0; i < config.typeDefinitions.length; i++) {
      var typeDef = config.typeDefinitions[i];
      if (!typeDef.id || typeDef.id.trim() === '') {
        errors.push('Type definition #' + (i + 1) + ' missing ID');
      }
      if (!typeDef.name || typeDef.name.trim() === '') {
        errors.push('Type definition #' + (i + 1) + ' missing name');
      }
      if (!typeDef.keywords || typeDef.keywords.length === 0) {
        errors.push('Type definition #' + (i + 1) + ' must have at least one keyword');
      }
    }
  }

  // Validate layout config
  if (config.layoutConfig) {
    if (config.layoutConfig.rows && config.layoutConfig.rows < 1) {
      errors.push('Layout must have at least 1 row');
    }
    if (config.layoutConfig.positionsPerRow && config.layoutConfig.positionsPerRow < 1) {
      errors.push('Layout must have at least 1 position per row');
    }
  }

  return {
    valid: errors.length === 0,
    errors: errors
  };
}

// ============================================================================
// TERMINOLOGY HELPER FUNCTIONS
// ============================================================================

/**
 * Gets user-friendly terminology for display
 * @param {string} key - Terminology key
 * @return {string} User-configured term
 */
function getTerminology(key) {
  var primaryEntity = getPrimaryEntityType();
  var layoutConfig = getLayoutConfig();

  var terminologyMap = {
    'entity_singular': primaryEntity.singular,
    'entity_plural': primaryEntity.plural,
    'entity_verb': primaryEntity.verb,
    'entity_singular_lower': primaryEntity.singular.toLowerCase(),
    'entity_plural_lower': primaryEntity.plural.toLowerCase(),
    'hierarchy_level_0': getHierarchyLevelName(0),
    'hierarchy_level_1': getHierarchyLevelName(1),
    'hierarchy_level_2': getHierarchyLevelName(2),
    'position_prefix': layoutConfig.labelPrefix
  };

  return terminologyMap[key] || key;
}

/**
 * Replaces terminology placeholders in a string
 * @param {string} template - String with placeholders like {{entity_singular}}
 * @return {string} String with replacements
 */
function replacePlaceholders(template) {
  if (!template) return '';

  return template.replace(/\{\{(\w+)\}\}/g, function(match, key) {
    return getTerminology(key);
  });
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Checks if system has been initialized
 * @return {boolean} True if initialized
 */
function isSystemInitialized() {
  var initialized = PropertiesService.getScriptProperties().getProperty(CONFIG_KEYS.SYSTEM_INITIALIZED);
  return initialized === 'true';
}

/**
 * Resets type system configuration (for testing or fresh start)
 * @return {Object} Reset result
 */
function resetTypeSystemConfiguration() {
  try {
    var scriptProps = PropertiesService.getScriptProperties();

    // Delete all configuration keys
    scriptProps.deleteProperty(CONFIG_KEYS.SYSTEM_INITIALIZED);
    scriptProps.deleteProperty(CONFIG_KEYS.PRIMARY_ENTITY_TYPE);
    scriptProps.deleteProperty(CONFIG_KEYS.TYPE_DEFINITIONS);
    scriptProps.deleteProperty(CONFIG_KEYS.CATEGORY_CLASSIFICATIONS);
    scriptProps.deleteProperty(CONFIG_KEYS.LAYOUT_CONFIG);
    scriptProps.deleteProperty(CONFIG_KEYS.HIERARCHY_LEVELS);
    scriptProps.deleteProperty(CONFIG_KEYS.MIGRATED_FROM_V1);

    Logger.log('Type system configuration reset successfully');

    return {
      success: true,
      message: 'Configuration reset successfully'
    };

  } catch (e) {
    Logger.log('Error resetting type system configuration: ' + e.message);
    return {
      success: false,
      message: 'Error resetting configuration: ' + e.message
    };
  }
}
