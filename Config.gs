/**
 * Configuration Constants for PTC Arena Sheets Generic Integration
 * Centralized configuration for sheet structure, colors, and mappings
 *
 * NOTE: Many configuration values are now dynamically loaded from TypeSystemConfig.gs
 * This provides user-configurable terminology and classifications.
 */

// DEPRECATED: Legacy sheet names for backward compatibility
// New installations use dynamic configuration from TypeSystemConfig.gs
var SHEET_NAMES = {
  LEGEND_NET: 'Legend-NET',
  OVERHEAD: 'Overhead',
  FULL_RACK_A: 'Full Rack Item A',
  FULL_RACK_B: 'Full Rack Item B',
  FULL_RACK_C: 'Full Rack Item C',
  FULL_RACK_D: 'Full Rack Item D',
  FULL_RACK_E: 'Full Rack Item E',
  FULL_RACK_F: 'Full Rack F',
  FULL_RACK_G: 'Full Rack Item G'
};

// Rack Tab Column Headers
var RACK_COLUMNS = {
  QTY: 'Qty',
  ITEM_NUMBER: 'Item Number',
  ITEM_NAME: 'Item Name',
  ITEM_CATEGORY: 'Item Category'
};

// Column indices (0-based) for rack tabs
var RACK_COLUMN_INDICES = {
  QTY: 0,        // Column A
  ITEM_NUMBER: 1, // Column B
  ITEM_NAME: 2,   // Column C
  ITEM_CATEGORY: 3 // Column D
};

// DEPRECATED: Legacy color schemes for backward compatibility
// New installations use dynamic configuration from TypeSystemConfig.gs
var RACK_COLORS = {
  FULL_RACK_A: '#00FFFF',    // Cyan
  FULL_RACK_B: '#FFA500',    // Orange
  FULL_RACK_C: '#00FF00',    // Green
  FULL_RACK_D: '#90EE90',    // Light Green
  FULL_RACK_E: '#90EE90',    // Light Green
  FULL_RACK_F: '#0000FF',    // Blue
  FULL_RACK_G: '#000000',    // Black
  ETH: '#FFFF00',            // Yellow
  SPINE: '#FF00FF',          // Purple/Magenta
  GRID_POD: '#00FF00'        // Green
};

// DEPRECATED: Legacy category definitions for backward compatibility
// New installations use dynamic configuration from TypeSystemConfig.gs
var LEGEND_CATEGORIES = {
  ETH: {
    name: 'Data center ETH',
    color: RACK_COLORS.ETH,
    prefix: 'ETH'
  },
  SPINE: {
    name: 'Data Center SPINE RACK',
    color: RACK_COLORS.SPINE,
    prefix: 'SPINE'
  },
  GRID_POD: {
    name: 'DATA CENTER GRID-POD',
    color: RACK_COLORS.GRID_POD,
    prefix: 'GRID'
  }
};

// Arena API field mappings
// Maps Arena API response fields to our sheet columns
var ARENA_FIELD_MAPPING = {
  // Arena field name → Sheet column purpose
  NUMBER: 'itemNumber',        // Arena part number
  NAME: 'itemName',            // Arena item name/description
  CATEGORY: 'category',        // Arena category field
  QUANTITY: 'quantity',        // Default quantity (may be overridden)
  DESCRIPTION: 'description',  // Detailed description
  REVISION: 'revisionNumber',  // Revision info
  LIFECYCLE: 'lifecyclePhase'  // Lifecycle status
};

// DEPRECATED: Legacy rack type and category keywords for backward compatibility
// New installations use dynamic configuration from TypeSystemConfig.gs
// These are kept only for reference and migration purposes
var RACK_TYPE_KEYWORDS = {
  FULL_RACK_A: ['rack a', 'type a', 'config a'],
  FULL_RACK_B: ['rack b', 'type b', 'config b'],
  FULL_RACK_C: ['rack c', 'type c', 'config c'],
  FULL_RACK_D: ['rack d', 'type d', 'config d'],
  FULL_RACK_E: ['rack e', 'type e', 'config e'],
  FULL_RACK_F: ['rack f', 'type f', 'config f'],
  FULL_RACK_G: ['rack g', 'type g', 'config g']
};

var CATEGORY_KEYWORDS = {
  ETH: ['ethernet', 'eth', 'network switch', 'nic'],
  SPINE: ['spine', 'backbone', 'core switch'],
  GRID_POD: ['grid', 'pod', 'power distribution'],
  CATEGORY_A: ['server', 'compute', 'cpu'],
  CATEGORY_B: ['storage', 'disk', 'ssd'],
  CATEGORY_C: ['networking', 'switch', 'router'],
  CATEGORY_D: ['power', 'pdu', 'ups'],
  CATEGORY_F: ['cable', 'fiber', 'copper'],
  CATEGORY_L: ['licensing', 'software']
};

// Header row styling
var HEADER_STYLE = {
  BACKGROUND_COLOR: '#4A86E8',  // Blue
  FONT_COLOR: '#FFFFFF',         // White
  FONT_WEIGHT: 'bold',
  FONT_SIZE: 11,
  HORIZONTAL_ALIGNMENT: 'center',
  VERTICAL_ALIGNMENT: 'middle'
};

// Legend-NET header styling
var LEGEND_HEADER_STYLE = {
  BACKGROUND_COLOR: '#0000FF',  // Blue
  FONT_COLOR: '#FFFFFF',         // White
  FONT_WEIGHT: 'bold',
  FONT_SIZE: 12,
  HORIZONTAL_ALIGNMENT: 'center'
};

// Totals row styling
var TOTALS_STYLE = {
  BACKGROUND_COLOR: '#0000FF',  // Blue
  FONT_COLOR: '#FFFFFF',         // White
  FONT_WEIGHT: 'bold',
  HORIZONTAL_ALIGNMENT: 'center'
};

// Overhead grid configuration
var OVERHEAD_CONFIG = {
  START_ROW: 2,              // Row where grid starts
  START_COL: 2,              // Column where grid starts (B)
  ROWS_PER_AISLE: 2,         // Number of rows per aisle
  POSITIONS_PER_ROW: 10,     // Number of rack positions per row
  CELL_HEIGHT: 25,           // Height of each cell
  CELL_WIDTH: 120            // Width of each cell
};

/**
 * Gets the color for a given entity/rack type
 * Now uses dynamic configuration from TypeSystemConfig.gs
 * @param {string} entityType - The entity type name
 * @return {string} Color hex code
 */
function getRackColor(entityType) {
  // Try to get color from dynamic configuration
  var typeDefinitions = getTypeDefinitions();
  for (var i = 0; i < typeDefinitions.length; i++) {
    var typeDef = typeDefinitions[i];
    if (typeDef.name === entityType || typeDef.displayName === entityType) {
      return typeDef.color;
    }
  }

  // Fallback to legacy hardcoded colors for backward compatibility
  var normalizedType = entityType.toUpperCase().replace(/\s+/g, '_');
  return RACK_COLORS[normalizedType] || '#CCCCCC'; // Default to gray
}

/**
 * Gets the category for an item based on keywords
 * REFACTORED: Now uses dynamic configuration from TypeSystemConfig.gs
 * @param {string} itemName - The item name or description
 * @param {string} itemCategory - The Arena category field
 * @return {string} Matched category
 */
function getCategoryFromItem(itemName, itemCategory) {
  var searchText = (itemName + ' ' + itemCategory).toLowerCase();

  // Use dynamic category classifications
  var categoryClassifications = getCategoryClassifications();

  for (var i = 0; i < categoryClassifications.length; i++) {
    var catClass = categoryClassifications[i];
    if (!catClass.enabled) continue;

    var keywords = catClass.keywords;
    for (var j = 0; j < keywords.length; j++) {
      if (searchText.indexOf(keywords[j].toLowerCase()) !== -1) {
        return catClass.name;
      }
    }
  }

  // If no category classifications configured, fall back to legacy
  if (categoryClassifications.length === 0) {
    for (var category in CATEGORY_KEYWORDS) {
      var legacyKeywords = CATEGORY_KEYWORDS[category];
      for (var k = 0; k < legacyKeywords.length; k++) {
        if (searchText.indexOf(legacyKeywords[k].toLowerCase()) !== -1) {
          return category.replace('_', ' ');
        }
      }
    }
  }

  return 'Uncategorized';
}

/**
 * Determines which entity type tab an item belongs to
 * REFACTORED: Now uses dynamic configuration from TypeSystemConfig.gs
 * @param {Object} arenaItem - The Arena API item object
 * @return {string} The entity type tab name, or null if no match
 */
function determineEntityType(arenaItem) {
  var searchText = '';

  // Build search text from various fields
  if (arenaItem.name) searchText += arenaItem.name.toLowerCase() + ' ';
  if (arenaItem.category) searchText += arenaItem.category.toLowerCase() + ' ';
  if (arenaItem.description) searchText += arenaItem.description.toLowerCase() + ' ';

  // Use dynamic type definitions
  var typeDefinitions = getTypeDefinitions();

  for (var i = 0; i < typeDefinitions.length; i++) {
    var typeDef = typeDefinitions[i];
    if (!typeDef.enabled) continue;

    var keywords = typeDef.keywords;
    for (var j = 0; j < keywords.length; j++) {
      if (searchText.indexOf(keywords[j].toLowerCase()) !== -1) {
        return typeDef.name; // Return the user-defined type name
      }
    }
  }

  return null; // No entity type match found
}

/**
 * @deprecated Use determineEntityType() instead
 * Kept for backward compatibility with existing code
 */
function determineRackType(arenaItem) {
  Logger.log('DEPRECATED: determineRackType() called. Use determineEntityType() instead.');
  return determineEntityType(arenaItem);
}

/**
 * Gets all entity type tab names as an array
 * REFACTORED: Now uses dynamic configuration from TypeSystemConfig.gs
 * @return {Array<string>} Array of entity tab names
 */
function getAllEntityTabNames() {
  var typeDefinitions = getTypeDefinitions();
  var tabNames = [];

  for (var i = 0; i < typeDefinitions.length; i++) {
    if (typeDefinitions[i].enabled) {
      tabNames.push(typeDefinitions[i].name);
    }
  }

  return tabNames;
}

/**
 * @deprecated Use getAllEntityTabNames() instead
 * Kept for backward compatibility with existing code
 */
function getAllRackTabNames() {
  Logger.log('DEPRECATED: getAllRackTabNames() called. Use getAllEntityTabNames() instead.');
  return getAllEntityTabNames();
}
