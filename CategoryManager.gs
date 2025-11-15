/**
 * Category Manager
 * Handles category hierarchy, colors, and BOM level configuration
 */

// Script Properties keys for category configuration
var PROP_CATEGORY_COLORS = 'CATEGORY_COLORS';
var PROP_BOM_HIERARCHY = 'BOM_HIERARCHY';
var PROP_ITEM_COLUMNS = 'ITEM_COLUMNS';
var PROP_FAVORITE_CATEGORIES = 'FAVORITE_CATEGORIES';

/**
 * Gets category color configuration
 * @return {Object} Map of category name/GUID to color
 */
function getCategoryColors() {
  var json = PropertiesService.getScriptProperties().getProperty(PROP_CATEGORY_COLORS);
  if (!json) {
    return getDefaultCategoryColors();
  }
  try {
    return JSON.parse(json);
  } catch (e) {
    return getDefaultCategoryColors();
  }
}

/**
 * Default category colors for datacenter items
 * @return {Object} Default color mappings
 */
function getDefaultCategoryColors() {
  return {
    'Hall': '#E8F5E9',           // Light green
    'Pod': '#FFF9C4',             // Light yellow
    'Rack': '#B3E5FC',            // Light blue
    'Server': '#F8BBD0',          // Light pink
    'Networking': '#FFECB3',      // Light amber
    'Storage': '#D1C4E9',         // Light purple
    'Power': '#FFE0B2',           // Light orange
    'Cable': '#C8E6C9',           // Light green
    'Component': '#F5F5F5',       // Light grey
    'default': '#FFFFFF'          // White
  };
}

/**
 * Saves category color configuration
 * @param {Object} colors - Map of category to color
 */
function saveCategoryColors(colors) {
  PropertiesService.getScriptProperties().setProperty(
    PROP_CATEGORY_COLORS,
    JSON.stringify(colors)
  );
}

/**
 * Gets BOM hierarchy configuration
 * @return {Array<Object>} Array of hierarchy levels
 */
function getBOMHierarchy() {
  var json = PropertiesService.getScriptProperties().getProperty(PROP_BOM_HIERARCHY);
  if (!json) {
    return getDefaultBOMHierarchy();
  }
  try {
    return JSON.parse(json);
  } catch (e) {
    return getDefaultBOMHierarchy();
  }
}

/**
 * Default BOM hierarchy for datacenters
 * @return {Array<Object>} Default hierarchy levels
 */
function getDefaultBOMHierarchy() {
  return [
    { level: 0, category: 'Hall', name: 'Hall' },
    { level: 1, category: 'Pod', name: 'Pod' },
    { level: 2, category: 'Rack', name: 'Rack' },
    { level: 3, category: 'Server', name: 'Server' },
    { level: 4, category: 'Component', name: 'Component' }
  ];
}

/**
 * Saves BOM hierarchy configuration
 * @param {Array<Object>} hierarchy - Hierarchy levels
 */
function saveBOMHierarchy(hierarchy) {
  PropertiesService.getScriptProperties().setProperty(
    PROP_BOM_HIERARCHY,
    JSON.stringify(hierarchy)
  );
}

/**
 * Gets configured item columns (attributes to show)
 * @return {Array<Object>} Array of column configurations
 */
function getItemColumns() {
  var json = PropertiesService.getScriptProperties().getProperty(PROP_ITEM_COLUMNS);
  if (!json) {
    return getDefaultItemColumns();
  }
  try {
    return JSON.parse(json);
  } catch (e) {
    return getDefaultItemColumns();
  }
}

/**
 * Default item columns for datacenter items
 * @return {Array<Object>} Default column configurations
 */
function getDefaultItemColumns() {
  return [
    { attributeGuid: null, attributeName: 'description', header: 'Description', width: 300 },
    { attributeGuid: null, attributeName: 'category', header: 'Category', width: 150 },
    { attributeGuid: null, attributeName: 'lifecyclePhase', header: 'Lifecycle', width: 100 }
  ];
}

/**
 * Saves item column configuration
 * @param {Array<Object>} columns - Column configurations
 */
function saveItemColumns(columns) {
  PropertiesService.getScriptProperties().setProperty(
    PROP_ITEM_COLUMNS,
    JSON.stringify(columns)
  );
}

/**
 * Gets all categories from Arena
 * @return {Array<Object>} Array of category objects
 */
function getArenaCategories() {
  try {
    var client = new ArenaAPIClient();
    var response = client.makeRequest('/settings/items/categories', { method: 'GET' });

    var categories = response.results || response.Results || [];

    return categories.map(function(cat) {
      return {
        guid: cat.guid || cat.Guid,
        name: cat.name || cat.Name,
        path: cat.path || cat.Path || '',
        fullPath: (cat.path || cat.Path || '') + '/' + (cat.name || cat.Name)
      };
    });
  } catch (error) {
    Logger.log('Error fetching categories: ' + error.message);
    return [];
  }
}

/**
 * Gets all item attributes from Arena
 * @return {Array<Object>} Array of attribute objects
 */
function getArenaAttributes() {
  try {
    var client = new ArenaAPIClient();
    var response = client.makeRequest('/settings/items/attributes', { method: 'GET' });

    var attributes = response.results || response.Results || [];

    return attributes.map(function(attr) {
      return {
        guid: attr.guid || attr.Guid,
        name: attr.name || attr.Name,
        apiName: attr.apiName || attr.ApiName,
        type: attr.type || attr.Type || 'SINGLE_LINE_TEXT',
        path: attr.path || attr.Path || ''
      };
    });
  } catch (error) {
    Logger.log('Error fetching attributes: ' + error.message);
    return [];
  }
}

/**
 * Gets color for a specific category
 * @param {string} categoryName - Category name or path
 * @return {string} Hex color code
 */
function getCategoryColor(categoryName) {
  var colors = getCategoryColors();

  // Try exact match first
  if (colors[categoryName]) {
    return colors[categoryName];
  }

  // Try partial match
  for (var key in colors) {
    if (categoryName.toLowerCase().indexOf(key.toLowerCase()) !== -1) {
      return colors[key];
    }
  }

  return colors.default || '#FFFFFF';
}

/**
 * Gets BOM level for a category
 * @param {string} categoryName - Category name
 * @return {number} BOM level (0-based)
 */
function getCategoryLevel(categoryName) {
  var hierarchy = getBOMHierarchy();

  for (var i = 0; i < hierarchy.length; i++) {
    if (hierarchy[i].category === categoryName) {
      return hierarchy[i].level;
    }
  }

  // Default to highest level if not found
  return 99;
}

/**
 * Determines parent category for a given category
 * @param {string} categoryName - Category name
 * @return {string|null} Parent category name or null
 */
function getParentCategory(categoryName) {
  var hierarchy = getBOMHierarchy();
  var currentLevel = getCategoryLevel(categoryName);

  if (currentLevel === 0) {
    return null; // Top level has no parent
  }

  // Find category at level - 1
  for (var i = 0; i < hierarchy.length; i++) {
    if (hierarchy[i].level === currentLevel - 1) {
      return hierarchy[i].category;
    }
  }

  return null;
}

/**
 * Validates category hierarchy configuration
 * @param {Array<Object>} hierarchy - Hierarchy to validate
 * @return {Object} Validation result
 */
function validateBOMHierarchy(hierarchy) {
  var errors = [];

  if (!Array.isArray(hierarchy)) {
    return { valid: false, errors: ['Hierarchy must be an array'] };
  }

  // Check for duplicate levels
  var levels = {};
  hierarchy.forEach(function(item) {
    if (levels[item.level]) {
      errors.push('Duplicate level: ' + item.level);
    }
    levels[item.level] = true;
  });

  // Check for gaps in levels
  var maxLevel = Math.max.apply(null, hierarchy.map(function(h) { return h.level; }));
  for (var i = 0; i <= maxLevel; i++) {
    if (!levels[i]) {
      errors.push('Missing level: ' + i);
    }
  }

  return {
    valid: errors.length === 0,
    errors: errors
  };
}

/**
 * Gets items for a specific category with lifecycle filtering
 * @param {string} categoryGuid - Category GUID
 * @param {string} lifecyclePhase - Lifecycle phase to filter (optional)
 * @return {Array<Object>} Array of items
 */
function getItemsByCategory(categoryGuid, lifecyclePhase) {
  try {
    var client = new ArenaAPIClient();

    // Build search criteria
    var criteria = [[{
      attribute: 'category.guid',
      operator: 'IS_EQUAL_TO',
      value: categoryGuid
    }]];

    // Add lifecycle filter if specified
    if (lifecyclePhase) {
      criteria[0].push({
        attribute: 'lifecyclePhase.name',
        operator: 'IS_EQUAL_TO',
        value: lifecyclePhase
      });
    }

    var criteriaParam = encodeURIComponent(JSON.stringify(criteria));
    var response = client.makeRequest('/items?criteria=' + criteriaParam, { method: 'GET' });

    return response.results || response.Results || [];
  } catch (error) {
    Logger.log('Error fetching items by category: ' + error.message);
    return [];
  }
}

/**
 * Searches items across all categories
 * @param {string} searchQuery - Search query
 * @param {string} lifecyclePhase - Lifecycle phase filter (optional)
 * @return {Array<Object>} Array of matching items
 */
function searchItems(searchQuery, lifecyclePhase) {
  try {
    var client = new ArenaAPIClient();

    var queryParams = ['searchQuery=' + encodeURIComponent(searchQuery)];

    if (lifecyclePhase) {
      queryParams.push('lifecyclePhase=' + encodeURIComponent(lifecyclePhase));
    }

    var endpoint = '/items/searches?' + queryParams.join('&');
    var response = client.makeRequest(endpoint, { method: 'GET' });

    return response.results || response.Results || [];
  } catch (error) {
    Logger.log('Error searching items: ' + error.message);
    return [];
  }
}

/**
 * Gets lifecycle phases from Arena
 * @return {Array<string>} Array of lifecycle phase names
 */
function getLifecyclePhases() {
  try {
    var client = new ArenaAPIClient();
    var response = client.makeRequest('/settings/items/lifecyclephases', { method: 'GET' });

    var phases = response.results || response.Results || [];

    return phases.map(function(phase) {
      return phase.name || phase.Name || '';
    });
  } catch (error) {
    Logger.log('Error fetching lifecycle phases: ' + error.message);
    return ['Production', 'Prototype', 'In Development', 'Obsolete'];
  }
}

/**
 * Gets an attribute value from an item
 * @param {Object} item - Item object
 * @param {string} attributeGuid - Attribute GUID to retrieve
 * @return {string|null} Attribute value or null if not found
 */
function getAttributeValue(item, attributeGuid) {
  if (!item || !attributeGuid) {
    return null;
  }

  // Check if the item has an attributes collection
  if (item.attributes || item.Attributes) {
    var attrs = item.attributes || item.Attributes;

    for (var i = 0; i < attrs.length; i++) {
      var attr = attrs[i];
      var guid = attr.guid || attr.Guid;

      if (guid === attributeGuid) {
        return attr.value || attr.Value || attr.displayValue || attr.DisplayValue || '';
      }
    }
  }

  // Fallback: check if attribute is a direct property
  // Some standard fields might be at root level
  var builtInFields = {
    'number': item.number || item.Number,
    'name': item.name || item.Name,
    'description': item.description || item.Description,
    'revisionNumber': item.revisionNumber || item.RevisionNumber,
    'lifecyclePhase': item.lifecyclePhase ? (item.lifecyclePhase.name || item.lifecyclePhase.Name) : null,
    'category': item.category ? (item.category.name || item.category.Name) : null
  };

  // Try to match by GUID or common field name
  for (var key in builtInFields) {
    if (builtInFields[key]) {
      return builtInFields[key];
    }
  }

  return null;
}
