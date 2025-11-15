# PTC Arena Sheets Data Center

Google Sheets Add-on for integrating with PTC Arena API.

## Features

- **Complete Arena API Authorization** with workspace ID support
- Secure credential storage using Google Apps Script Properties Service
- User-friendly configuration wizard
- Connection testing and validation
- Workspace-scoped API requests

## Configuration

The add-on requires three pieces of information to connect to Arena API:

1. **API Endpoint** - The base URL for your Arena API instance (e.g., `https://api.arenasolutions.com`)
2. **Workspace ID** - Your Arena workspace identifier
3. **API Key** - Your Arena API authentication key

### How to Configure

1. Open your Google Sheet
2. Click on **Arena Data Center** menu → **Configure Arena Connection**
3. Fill in all three required fields:
   - API Endpoint
   - Workspace ID
   - API Key
4. Click **Save & Test Connection**
5. The wizard will automatically test the connection and close if successful

## Files

### Code.gs
Main entry point for the add-on. Contains:
- Menu initialization
- Login wizard display
- Connection testing
- Credential management UI

### Authorization.gs
Handles credential storage and retrieval. Manages:
- API Endpoint
- **Workspace ID** (required for all Arena API operations)
- API Key
- Credential validation

### LoginWizard.html
User interface for credential configuration. Features:
- Form validation
- Visual feedback
- Automatic connection testing
- Current configuration loading

### ArenaAPI.gs
Arena API client implementation. Includes:
- Authenticated API requests with workspace ID in headers
- Workspace-scoped endpoint construction
- Error handling
- Helper methods for common operations:
  - `testConnection()` - Verify API connectivity
  - `getItems()` - Get items from workspace
  - `getItem(itemId)` - Get specific item
  - `createItem(itemData)` - Create new item
  - `updateItem(itemId, itemData)` - Update existing item
  - `getWorkspaceInfo()` - Get workspace details

## Workspace ID Usage

The workspace ID is used in two ways:

1. **HTTP Header**: Sent as `X-Arena-Workspace-Id` in all API requests
2. **URL Path**: Included in API endpoints (e.g., `/api/workspaces/{workspaceId}/items`)

This ensures all API operations are properly scoped to the correct Arena workspace.

## Security

- Credentials are stored using `PropertiesService.getUserProperties()`, which is user-specific and encrypted
- API keys are never logged or displayed in the UI after initial entry
- Credentials can be cleared at any time via the menu

## Development

This is a Google Apps Script project. To deploy:

1. Create a new Google Sheets document
2. Open **Extensions** → **Apps Script**
3. Copy the contents of each `.gs` and `.html` file to the script editor
4. Save and reload the spreadsheet
5. The **Arena Data Center** menu will appear

## API Endpoint Examples

The client includes workspace-aware API methods:

```javascript
var client = new ArenaAPIClient();

// Get all items in the workspace
var items = client.getItems();

// Get a specific item
var item = client.getItem('ITEM-12345');

// Create a new item
var newItem = client.createItem({
  name: 'New Part',
  description: 'Part description'
});

// Get workspace information
var workspace = client.getWorkspaceInfo();
```

All requests automatically include the configured workspace ID.

## Future Enhancements

- Data import/export functionality
- Batch operations
- Custom field mapping
- Automated sync schedules
- Advanced filtering and querying
