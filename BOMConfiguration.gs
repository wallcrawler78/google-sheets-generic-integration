/**
 * BOM Configuration Management
 * Handles configuration for BOM-specific attributes like position tracking
 */

// Property keys for BOM configuration storage
var BOM_CONFIG_KEYS = {
  POSITION_ATTRIBUTE_GUID: 'BOM_POSITION_ATTRIBUTE_GUID',
  POSITION_ATTRIBUTE_NAME: 'BOM_POSITION_ATTRIBUTE_NAME',
  POSITION_FORMAT: 'BOM_POSITION_FORMAT'
};

/**
 * Shows the Rack BOM Location Setting configuration dialog
 */
function showRackBOMLocationSetting() {
  var html = HtmlService.createHtmlOutputFromFile('ConfigureBOMPositionAttribute')
    .setWidth(600)
    .setHeight(550)
    .setTitle('Rack BOM Location Setting');
  SpreadsheetApp.getUi().showModalDialog(html, 'Rack BOM Location Configuration');
}

/**
 * Loads BOM position attribute configuration data for the UI
 * @return {Object} Configuration data including available attributes and current selection
 */
function loadBOMPositionConfigData() {
  try {
    var client = new ArenaAPIClient();

    // Get all BOM attributes from Arena
    // Fetches BOM attributes (attributes that can be set on BOM lines)
    var bomAttributes = getBOMAttributes(client);

    // Filter to text-type attributes only (exclude checkboxes, numbers, etc.)
    var textAttributes = bomAttributes.filter(function(attr) {
      return attr.fieldType === 'SINGLE_LINE_TEXT' ||
             attr.fieldType === 'MULTI_LINE_TEXT' ||
             attr.fieldType === 'FIXED_DROP_DOWN';
    });

    // Get current configuration
    var userProperties = PropertiesService.getUserProperties();
    var currentAttributeGuid = userProperties.getProperty(BOM_CONFIG_KEYS.POSITION_ATTRIBUTE_GUID);
    var currentAttributeName = userProperties.getProperty(BOM_CONFIG_KEYS.POSITION_ATTRIBUTE_NAME);
    var positionFormat = userProperties.getProperty(BOM_CONFIG_KEYS.POSITION_FORMAT) || 'Pos {n}';

    // Validate that the stored attribute still exists in BOM attributes
    var isValid = false;
    var validationWarning = null;

    if (currentAttributeGuid) {
      isValid = textAttributes.some(function(attr) {
        return attr.guid === currentAttributeGuid;
      });

      if (!isValid) {
        validationWarning = 'Previously configured attribute "' + (currentAttributeName || currentAttributeGuid) +
                          '" not found in current BOM attributes. It may have been deleted or is not a valid BOM attribute. ' +
                          'Please select a valid BOM-level attribute.';
        Logger.log('⚠ Configuration validation: ' + validationWarning);
        // Clear invalid configuration
        currentAttributeGuid = null;
        currentAttributeName = null;
      }
    }

    return {
      availableAttributes: textAttributes,
      currentSelection: {
        guid: currentAttributeGuid,
        name: currentAttributeName,
        format: positionFormat
      },
      validationWarning: validationWarning
    };

  } catch (error) {
    Logger.log('Error loading BOM position config data: ' + error.message);
    throw error;
  }
}

/**
 * Gets BOM-level attributes from Arena
 * @param {ArenaAPIClient} client - Initialized Arena API client
 * @return {Array} Array of BOM attribute objects
 */
function getBOMAttributes(client) {
  try {
    // Arena API endpoint for BOM attributes
    // Note: This might need adjustment based on actual Arena API structure
    // For now, we'll use the item attributes endpoint as a placeholder

    // Get attributes from workspace settings
    var endpoint = '/settings/items/attributes';
    var response = client.makeRequest(endpoint);

    if (response && response.results) {
      return response.results.map(function(attr) {
        return {
          guid: attr.guid,
          name: attr.name,
          fieldType: attr.fieldType,
          apiName: attr.apiName,
          fullPath: attr.category ? (attr.category + ' > ' + attr.name) : attr.name
        };
      });
    }

    return [];

  } catch (error) {
    Logger.log('Error fetching BOM attributes: ' + error.message);
    // Return empty array instead of throwing - allows UI to still work
    return [];
  }
}

/**
 * Validates a BOM attribute configuration by checking if it exists in Arena
 * @param {Object} config - Configuration object with guid and name
 * @return {Object} Validation result with success flag and message
 */
function validateBOMAttributeConfig(config) {
  try {
    if (!config || !config.guid) {
      return {
        success: false,
        message: 'No attribute selected. Please select a BOM attribute to validate.'
      };
    }

    Logger.log('Validating BOM attribute config: ' + config.name + ' (' + config.guid + ')');

    var client = new ArenaAPIClient();

    // 1. Fetch current BOM attributes from Arena
    var bomAttributes = getBOMAttributes(client);

    if (!bomAttributes || bomAttributes.length === 0) {
      return {
        success: false,
        message: 'Could not fetch BOM attributes from Arena. Check your connection.'
      };
    }

    // 2. Check if the selected attribute exists in the BOM attributes list
    var attributeExists = false;
    var attributeDetails = null;

    for (var i = 0; i < bomAttributes.length; i++) {
      if (bomAttributes[i].guid === config.guid) {
        attributeExists = true;
        attributeDetails = bomAttributes[i];
        break;
      }
    }

    if (!attributeExists) {
      return {
        success: false,
        message: 'Attribute "' + config.name + '" not found in Arena BOM attributes. It may have been deleted or is an Item attribute (not a BOM attribute).'
      };
    }

    // 3. Check if it's a text type (should be filtered already, but double-check)
    var validTypes = ['SINGLE_LINE_TEXT', 'MULTI_LINE_TEXT', 'FIXED_DROP_DOWN'];
    if (validTypes.indexOf(attributeDetails.fieldType) === -1) {
      return {
        success: false,
        message: 'Attribute "' + config.name + '" is type "' + attributeDetails.fieldType + '" which cannot store position text. Use a text-type attribute.'
      };
    }

    // 4. All checks passed
    Logger.log('✓ Validation passed for BOM attribute: ' + config.name);

    return {
      success: true,
      message: 'Attribute "' + config.name + '" is a valid BOM attribute and can be used for position tracking.',
      attributeDetails: {
        guid: attributeDetails.guid,
        name: attributeDetails.name,
        fieldType: attributeDetails.fieldType,
        apiName: attributeDetails.apiName
      }
    };

  } catch (error) {
    Logger.log('Error validating BOM attribute config: ' + error.message);
    return {
      success: false,
      message: 'Validation error: ' + error.message
    };
  }
}

/**
 * Saves the selected BOM position attribute configuration
 * @param {Object} config - Configuration object with guid, name, and format
 * @return {Object} Success result
 */
function saveBOMPositionAttribute(config) {
  try {
    var userProperties = PropertiesService.getUserProperties();

    if (config.guid && config.name) {
      // Save the selected attribute
      userProperties.setProperty(BOM_CONFIG_KEYS.POSITION_ATTRIBUTE_GUID, config.guid);
      userProperties.setProperty(BOM_CONFIG_KEYS.POSITION_ATTRIBUTE_NAME, config.name);
      Logger.log('Saved BOM position attribute: ' + config.name);
    } else {
      // Clear configuration if nothing selected
      userProperties.deleteProperty(BOM_CONFIG_KEYS.POSITION_ATTRIBUTE_GUID);
      userProperties.deleteProperty(BOM_CONFIG_KEYS.POSITION_ATTRIBUTE_NAME);
      Logger.log('Cleared BOM position attribute configuration');
    }

    // Save position format
    if (config.format) {
      userProperties.setProperty(BOM_CONFIG_KEYS.POSITION_FORMAT, config.format);
    }

    return {
      success: true,
      message: 'BOM position attribute configuration saved successfully'
    };

  } catch (error) {
    Logger.log('Error saving BOM position attribute: ' + error.message);
    throw error;
  }
}

/**
 * Gets the configured BOM position attribute
 * @return {Object|null} Configuration object or null if not configured
 */
function getBOMPositionAttributeConfig() {
  var userProperties = PropertiesService.getUserProperties();

  var guid = userProperties.getProperty(BOM_CONFIG_KEYS.POSITION_ATTRIBUTE_GUID);
  var name = userProperties.getProperty(BOM_CONFIG_KEYS.POSITION_ATTRIBUTE_NAME);
  var format = userProperties.getProperty(BOM_CONFIG_KEYS.POSITION_FORMAT) || 'Pos {n}';

  if (!guid || !name) {
    return null;
  }

  return {
    guid: guid,
    name: name,
    format: format
  };
}

/**
 * Clears the BOM position attribute configuration
 */
function clearBOMPositionAttribute() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(BOM_CONFIG_KEYS.POSITION_ATTRIBUTE_GUID);
  userProperties.deleteProperty(BOM_CONFIG_KEYS.POSITION_ATTRIBUTE_NAME);
  userProperties.deleteProperty(BOM_CONFIG_KEYS.POSITION_FORMAT);
  Logger.log('BOM position attribute configuration cleared');
}
