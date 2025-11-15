/**
 * Arena API Client
 * Handles communication with the Arena API using configured credentials including workspace ID
 */

/**
 * Arena API Client Class
 */
var ArenaAPIClient = function() {
  var credentials = getArenaCredentials();

  if (!credentials) {
    throw new Error('Arena API credentials not configured. Please configure the connection first.');
  }

  this.apiEndpoint = credentials.apiEndpoint;
  this.apiKey = credentials.apiKey;
  this.workspaceId = credentials.workspaceId;
};

/**
 * Makes an authenticated request to the Arena API
 * @param {string} endpoint - The API endpoint path (will be appended to base URL)
 * @param {Object} options - Request options (method, payload, etc.)
 * @return {Object} Parsed JSON response
 */
ArenaAPIClient.prototype.makeRequest = function(endpoint, options) {
  options = options || {};

  var url = this.apiEndpoint + endpoint;

  // Build headers with API key authentication
  var headers = {
    'Authorization': 'Bearer ' + this.apiKey,
    'Content-Type': 'application/json',
    'X-Arena-Workspace-Id': this.workspaceId  // Include workspace ID in headers
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
        }
      } catch (e) {
        errorMessage += ': ' + responseText;
      }

      throw new Error(errorMessage);
    }
  } catch (error) {
    Logger.log('API request error: ' + error.message);
    throw error;
  }
};

/**
 * Tests the connection to the Arena API
 * @return {Object} Result object with success status
 */
ArenaAPIClient.prototype.testConnection = function() {
  try {
    // Try to make a simple API call to verify connectivity
    // This endpoint might need to be adjusted based on actual Arena API
    var result = this.makeRequest('/api/health', { method: 'GET' });

    return {
      success: true,
      message: 'Successfully connected to Arena API',
      workspaceId: this.workspaceId
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

  var endpoint = '/api/workspaces/' + this.workspaceId + '/items';

  // Add query parameters if provided
  if (options.category) {
    endpoint += '?category=' + encodeURIComponent(options.category);
  }

  return this.makeRequest(endpoint, { method: 'GET' });
};

/**
 * Gets a specific item by ID from Arena API
 * @param {string} itemId - The item identifier
 * @return {Object} Item data
 */
ArenaAPIClient.prototype.getItem = function(itemId) {
  var endpoint = '/api/workspaces/' + this.workspaceId + '/items/' + encodeURIComponent(itemId);
  return this.makeRequest(endpoint, { method: 'GET' });
};

/**
 * Creates a new item in Arena
 * @param {Object} itemData - The item data to create
 * @return {Object} Created item data
 */
ArenaAPIClient.prototype.createItem = function(itemData) {
  var endpoint = '/api/workspaces/' + this.workspaceId + '/items';
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
  var endpoint = '/api/workspaces/' + this.workspaceId + '/items/' + encodeURIComponent(itemId);
  return this.makeRequest(endpoint, {
    method: 'PUT',
    payload: itemData
  });
};

/**
 * Gets workspace information
 * @return {Object} Workspace data
 */
ArenaAPIClient.prototype.getWorkspaceInfo = function() {
  var endpoint = '/api/workspaces/' + this.workspaceId;
  return this.makeRequest(endpoint, { method: 'GET' });
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
