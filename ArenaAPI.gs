/**
 * Arena API Client
 * Handles communication with the Arena API using session-based authentication
 */

/**
 * Arena API Client Class
 */
var ArenaAPIClient = function() {
  var credentials = getArenaCredentials();

  if (!credentials) {
    throw new Error('Arena API credentials not configured. Please configure the connection first.');
  }

  this.apiBase = credentials.apiBase;
  this.workspaceId = credentials.workspaceId;

  // Get a valid session ID (will login if necessary)
  this.sessionId = getValidSessionId();
};

/**
 * Makes an authenticated request to the Arena API
 * @param {string} endpoint - The API endpoint path (will be appended to base URL)
 * @param {Object} options - Request options (method, payload, etc.)
 * @return {Object} Parsed JSON response
 */
ArenaAPIClient.prototype.makeRequest = function(endpoint, options) {
  options = options || {};

  var url = this.apiBase + endpoint;

  // Build headers with session-based authentication
  var headers = {
    'arena_session_id': this.sessionId,
    'Content-Type': 'application/json'
  };

  // Add any custom headers from options
  if (options.headers) {
    for (var key in options.headers) {
      headers[key] = options.headers[key];
    }
  }

  // Build request options
  var requestOptions = {
    method: options.method || 'GET',
    headers: headers,
    muteHttpExceptions: true
  };

  // Add payload for POST/PUT requests
  if (options.payload) {
    requestOptions.payload = JSON.stringify(options.payload);
  }

  try {
    Logger.log('Making request to: ' + url);
    var response = UrlFetchApp.fetch(url, requestOptions);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    Logger.log('Response code: ' + responseCode);

    // Check if session expired (401 unauthorized)
    if (responseCode === 401) {
      Logger.log('Session expired, attempting to re-login...');
      clearSession();
      this.sessionId = getValidSessionId();

      // Retry the request with new session
      headers['arena_session_id'] = this.sessionId;
      requestOptions.headers = headers;

      response = UrlFetchApp.fetch(url, requestOptions);
      responseCode = response.getResponseCode();
      responseText = response.getContentText();
    }

    if (responseCode >= 200 && responseCode < 300) {
      // Success - parse and return JSON
      if (responseText) {
        return JSON.parse(responseText);
      }
      return { success: true };
    } else {
      // Error response
      var errorMessage = 'HTTP ' + responseCode;
      try {
        var errorData = JSON.parse(responseText);
        if (errorData.message) {
          errorMessage += ': ' + errorData.message;
        } else if (errorData.error) {
          errorMessage += ': ' + errorData.error;
        } else if (errorData.errors) {
          errorMessage += ': ' + JSON.stringify(errorData.errors);
        } else {
          errorMessage += ': ' + JSON.stringify(errorData);
        }
      } catch (e) {
        if (responseText && responseText.length < 500) {
          errorMessage += ': ' + responseText;
        } else if (responseText) {
          errorMessage += ': ' + responseText.substring(0, 500) + '...';
        }
      }

      Logger.log('API Error - URL: ' + url);
      Logger.log('API Error - Code: ' + responseCode);
      Logger.log('API Error - Response: ' + responseText);

      throw new Error(errorMessage);
    }
  } catch (error) {
    Logger.log('API request error: ' + error.message);
    throw error;
  }
};

/**
 * Tests the connection to the Arena API
 * @return {Object} Result object with success status and metrics
 */
ArenaAPIClient.prototype.testConnection = function() {
  try {
    // If we got this far, the connection is working (constructor already logged in)
    // Now gather some metrics
    Logger.log('Testing connection and gathering metrics...');

    // Get all items to calculate metrics
    var items = this.getAllItems(400);

    // Count unique categories
    var categorySet = {};
    items.forEach(function(item) {
      var categoryObj = item.category || item.Category || {};
      var categoryName = categoryObj.name || categoryObj.Name || 'Uncategorized';
      categorySet[categoryName] = true;
    });

    var categoryCount = Object.keys(categorySet).length;

    Logger.log('Connection test successful - ' + items.length + ' items, ' + categoryCount + ' categories');

    return {
      success: true,
      message: 'Successfully connected to Arena API',
      workspaceId: this.workspaceId,
      totalItems: items.length,
      categoryCount: categoryCount
    };
  } catch (error) {
    Logger.log('Connection test failed: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
};

/**
 * Gets items from Arena API with workspace context
 * @param {Object} options - Query options (filters, pagination, etc.)
 * @return {Object} Items data from Arena
 */
ArenaAPIClient.prototype.getItems = function(options) {
  options = options || {};

  var endpoint = '/items';

  // Add query parameters if provided
  var queryParams = [];
  if (options.category) {
    queryParams.push('category=' + encodeURIComponent(options.category));
  }
  if (options.offset) {
    queryParams.push('offset=' + options.offset);
  }
  if (options.limit) {
    queryParams.push('limit=' + options.limit);
  }

  if (queryParams.length > 0) {
    endpoint += '?' + queryParams.join('&');
  }

  return this.makeRequest(endpoint, { method: 'GET' });
};

/**
 * Gets a specific item by ID from Arena API
 * @param {string} itemId - The item identifier
 * @return {Object} Item data
 */
ArenaAPIClient.prototype.getItem = function(itemId) {
  var endpoint = '/items/' + encodeURIComponent(itemId);
  return this.makeRequest(endpoint, { method: 'GET' });
};

/**
 * Creates a new item in Arena
 * @param {Object} itemData - The item data to create
 * @return {Object} Created item data
 */
ArenaAPIClient.prototype.createItem = function(itemData) {
  var endpoint = '/items';
  return this.makeRequest(endpoint, {
    method: 'POST',
    payload: itemData
  });
};

/**
 * Updates an existing item in Arena
 * @param {string} itemId - The item identifier
 * @param {Object} itemData - The updated item data
 * @return {Object} Updated item data
 */
ArenaAPIClient.prototype.updateItem = function(itemId, itemData) {
  var endpoint = '/items/' + encodeURIComponent(itemId);
  return this.makeRequest(endpoint, {
    method: 'PUT',
    payload: itemData
  });
};

/**
 * Sets a custom attribute value on an item in Arena
 * @param {string} itemId - The item identifier (GUID)
 * @param {string} attributeGuid - The attribute GUID
 * @param {string} value - The value to set
 * @return {Object} Updated item data
 */
ArenaAPIClient.prototype.setItemAttribute = function(itemId, attributeGuid, value) {
  var endpoint = '/items/' + encodeURIComponent(itemId);

  Logger.log('Setting attribute on item: ' + itemId);
  Logger.log('Attribute GUID: ' + attributeGuid);
  Logger.log('Value: ' + value);

  // Arena API uses "additionalAttributes" not "attributes"
  var payload = {
    additionalAttributes: [
      {
        guid: attributeGuid,
        value: value
      }
    ]
  };

  Logger.log('setItemAttribute payload: ' + JSON.stringify(payload));

  return this.makeRequest(endpoint, {
    method: 'PUT',
    payload: payload
  });
};

/**
 * Gets workspace information
 * @return {Object} Workspace data
 */
ArenaAPIClient.prototype.getWorkspaceInfo = function() {
  var endpoint = '/settings/workspace';
  return this.makeRequest(endpoint, { method: 'GET' });
};

/**
 * Gets items filtered by category
 * @param {string} category - Category to filter by
 * @return {Object} Filtered items data
 */
ArenaAPIClient.prototype.getItemsByCategory = function(category) {
  return this.getItems({ category: category });
};

/**
 * Searches for items matching a query
 * @param {string} query - Search query string
 * @param {Object} options - Additional search options
 * @return {Object} Search results
 */
ArenaAPIClient.prototype.searchItems = function(query, options) {
  options = options || {};

  var endpoint = '/items/searches';

  var queryParams = [];
  queryParams.push('searchQuery=' + encodeURIComponent(query));

  if (options.limit) {
    queryParams.push('limit=' + options.limit);
  }
  if (options.offset) {
    queryParams.push('offset=' + options.offset);
  }

  endpoint += '?' + queryParams.join('&');

  return this.makeRequest(endpoint, { method: 'GET' });
};

/**
 * Gets a specific item by item number
 * @param {string} itemNumber - The item number to search for
 * @return {Object|null} Item data or null if not found
 */
ArenaAPIClient.prototype.getItemByNumber = function(itemNumber) {
  try {
    Logger.log('Searching for item by number: ' + itemNumber);

    // Fetch all items and filter (searchItems endpoint is broken)
    // Use getAllItems() for comprehensive search
    Logger.log('Fetching all items to search for: ' + itemNumber);
    var allItems = this.getAllItems();

    Logger.log('Fetched ' + allItems.length + ' items, searching for exact match: ' + itemNumber);

    // Find exact match by item number
    for (var i = 0; i < allItems.length; i++) {
      var item = allItems[i];
      var number = item.number || item.Number || '';

      if (number === itemNumber) {
        Logger.log('✓ Found exact match for item: ' + itemNumber);
        return item;
      }
    }

    Logger.log('Item not found: ' + itemNumber);
    return null;

  } catch (error) {
    Logger.log('Error getting item by number: ' + error.message);
    throw error;
  }
};

/**
 * Gets detailed attributes for a specific item
 * @param {string} itemId - The item identifier
 * @return {Object} Detailed item attributes
 */
ArenaAPIClient.prototype.getItemAttributes = function(itemId) {
  var endpoint = '/items/' + encodeURIComponent(itemId) + '/attributes';
  return this.makeRequest(endpoint, { method: 'GET' });
};

/**
 * Gets multiple items by their IDs in bulk
 * @param {Array<string>} itemIds - Array of item identifiers
 * @return {Array<Object>} Array of item data
 */
ArenaAPIClient.prototype.getBulkItems = function(itemIds) {
  if (!itemIds || itemIds.length === 0) {
    return [];
  }

  // Arena API typically doesn't have a bulk endpoint, fetch individually
  var results = [];

  itemIds.forEach(function(itemId) {
    try {
      var item = this.getItem(itemId);
      results.push(item);
    } catch (itemError) {
      Logger.log('Error fetching item ' + itemId + ': ' + itemError.message);
    }
  }.bind(this));

  return results;
};

/**
 * Gets all items with pagination support
 * @param {number} batchSize - Number of items per request (default 100)
 * @return {Array<Object>} All items from the workspace
 */
ArenaAPIClient.prototype.getAllItems = function(batchSize) {
  batchSize = batchSize || 400; // Use max batch size (400) for better performance

  var allItems = [];
  var offset = 0;
  var hasMore = true;

  while (hasMore) {
    try {
      var response = this.getItems({ limit: batchSize, offset: offset });

      // Handle different possible response structures (Arena uses capital R)
      var items = response.results || response.Results || response.items || response.data || [];

      if (Array.isArray(response)) {
        items = response;
      }

      if (items.length > 0) {
        allItems = allItems.concat(items);
        offset += items.length;

        Logger.log('Fetched ' + items.length + ' items (total: ' + allItems.length + ')');

        // Check if there are more items
        if (items.length < batchSize) {
          hasMore = false;
        }
      } else {
        hasMore = false;
      }

      // Add a small delay to avoid rate limiting
      Utilities.sleep(200);

    } catch (error) {
      Logger.log('Error fetching items at offset ' + offset + ': ' + error.message);
      hasMore = false;
    }
  }

  Logger.log('Fetched ' + allItems.length + ' total items from Arena');
  return allItems;
};

/**
 * Gets items filtered by multiple criteria
 * @param {Object} filters - Filter criteria (category, status, etc.)
 * @return {Array<Object>} Filtered items
 */
ArenaAPIClient.prototype.getFilteredItems = function(filters) {
  var response = this.getItems(filters);

  // Extract items from response
  return response.results || response.items || response.data || response;
};

/**
 * Exports data from Arena to the current spreadsheet
 * This is a helper function that demonstrates workspace-scoped data retrieval
 * @param {string} dataType - Type of data to export
 */
ArenaAPIClient.prototype.exportToSheet = function(dataType) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  try {
    var data;

    switch (dataType) {
      case 'items':
        data = this.getItems();
        break;
      case 'workspace':
        data = this.getWorkspaceInfo();
        break;
      default:
        throw new Error('Unknown data type: ' + dataType);
    }

    // Clear existing content
    sheet.clear();

    // Add workspace ID header
    sheet.getRange(1, 1).setValue('Workspace ID:');
    sheet.getRange(1, 2).setValue(this.workspaceId);

    // Add data (implementation depends on actual API response structure)
    // This is a placeholder that would need to be customized
    sheet.getRange(3, 1).setValue('Data:');
    sheet.getRange(4, 1).setValue(JSON.stringify(data, null, 2));

    return { success: true, message: 'Data exported successfully' };

  } catch (error) {
    Logger.log('Export error: ' + error.message);
    throw error;
  }
};
