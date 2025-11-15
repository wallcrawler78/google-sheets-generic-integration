/**
 * Arena API Authorization Management
 * Handles storing and retrieving Arena API credentials including workspace ID
 */

// Property keys for credential storage
var PROPERTY_KEYS = {
  API_ENDPOINT: 'ARENA_API_ENDPOINT',
  API_KEY: 'ARENA_API_KEY',
  WORKSPACE_ID: 'ARENA_WORKSPACE_ID'
};

/**
 * Saves Arena API credentials to user properties
 * @param {Object} credentials - Object containing apiEndpoint, apiKey, and workspaceId
 * @return {Object} Result object with success status
 */
function saveArenaCredentials(credentials) {
  try {
    var userProperties = PropertiesService.getUserProperties();

    // Validate required fields
    if (!credentials.apiEndpoint || credentials.apiEndpoint.trim() === '') {
      throw new Error('API Endpoint is required');
    }
    if (!credentials.apiKey || credentials.apiKey.trim() === '') {
      throw new Error('API Key is required');
    }
    if (!credentials.workspaceId || credentials.workspaceId.trim() === '') {
      throw new Error('Workspace ID is required');
    }

    // Clean up the API endpoint (remove trailing slash)
    var cleanEndpoint = credentials.apiEndpoint.trim().replace(/\/$/, '');

    // Save credentials
    userProperties.setProperty(PROPERTY_KEYS.API_ENDPOINT, cleanEndpoint);
    userProperties.setProperty(PROPERTY_KEYS.API_KEY, credentials.apiKey.trim());
    userProperties.setProperty(PROPERTY_KEYS.WORKSPACE_ID, credentials.workspaceId.trim());

    Logger.log('Arena API credentials saved successfully');
    return { success: true, message: 'Credentials saved successfully' };

  } catch (error) {
    Logger.log('Error saving credentials: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Retrieves the Arena API endpoint
 * @return {string|null} The API endpoint or null if not set
 */
function getApiEndpoint() {
  return PropertiesService.getUserProperties().getProperty(PROPERTY_KEYS.API_ENDPOINT);
}

/**
 * Retrieves the Arena API key
 * @return {string|null} The API key or null if not set
 */
function getApiKey() {
  return PropertiesService.getUserProperties().getProperty(PROPERTY_KEYS.API_KEY);
}

/**
 * Retrieves the Arena workspace ID
 * @return {string|null} The workspace ID or null if not set
 */
function getWorkspaceId() {
  return PropertiesService.getUserProperties().getProperty(PROPERTY_KEYS.WORKSPACE_ID);
}

/**
 * Retrieves all Arena API credentials
 * @return {Object|null} Object containing all credentials or null if not configured
 */
function getArenaCredentials() {
  var apiEndpoint = getApiEndpoint();
  var apiKey = getApiKey();
  var workspaceId = getWorkspaceId();

  if (!apiEndpoint || !apiKey || !workspaceId) {
    return null;
  }

  return {
    apiEndpoint: apiEndpoint,
    apiKey: apiKey,
    workspaceId: workspaceId
  };
}

/**
 * Checks if Arena API credentials are configured
 * @return {boolean} True if all required credentials are set
 */
function isAuthorized() {
  var credentials = getArenaCredentials();
  return credentials !== null;
}

/**
 * Clears all stored Arena API credentials
 */
function clearArenaCredentials() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(PROPERTY_KEYS.API_ENDPOINT);
  userProperties.deleteProperty(PROPERTY_KEYS.API_KEY);
  userProperties.deleteProperty(PROPERTY_KEYS.WORKSPACE_ID);
  Logger.log('Arena API credentials cleared');
}

/**
 * Gets the current authorization status for display
 * @return {Object} Status object with configuration details
 */
function getAuthorizationStatus() {
  var credentials = getArenaCredentials();

  if (!credentials) {
    return {
      isConfigured: false,
      message: 'Not configured'
    };
  }

  return {
    isConfigured: true,
    apiEndpoint: credentials.apiEndpoint,
    workspaceId: credentials.workspaceId,
    // Don't return the actual API key for security
    hasApiKey: true
  };
}
