/**
 * Status Manager
 * Manages sync status indicators for rack configuration and overview sheets
 */

// ============================================================================
// STATUS CONSTANTS
// ============================================================================

/**
 * Status values for rack configuration sheets
 */
var RACK_STATUS = {
  PLACEHOLDER: 'PLACEHOLDER',      // Created locally, not yet in Arena
  SYNCED: 'SYNCED',               // Matches Arena exactly
  LOCAL_MODIFIED: 'LOCAL_MODIFIED', // User edited after sync
  ARENA_MODIFIED: 'ARENA_MODIFIED', // Arena changed externally
  ERROR: 'ERROR'                   // Last sync failed
};

/**
 * Status indicators (emoji with text fallback)
 * Using Unicode colored circles
 */
var STATUS_INDICATORS = {
  PLACEHOLDER: '\u{1F534}',      // 🔴 Red circle
  SYNCED: '\u{1F7E2}',           // 🟢 Green circle
  LOCAL_MODIFIED: '\u{1F7E0}',   // 🟠 Orange circle
  ARENA_MODIFIED: '\u{1F7E1}',   // 🟡 Yellow circle
  ERROR: '\u274C'                 // ❌ Cross mark
};

/**
 * Text fallback indicators (if emojis don't work)
 */
var STATUS_TEXT_INDICATORS = {
  PLACEHOLDER: '[NEW]',
  SYNCED: '[OK]',
  LOCAL_MODIFIED: '[EDIT]',
  ARENA_MODIFIED: '[DIFF]',
  ERROR: '[ERR]'
};

/**
 * DEPRECATED: Metadata column constants (no longer used in Row 1)
 * These columns have been moved to the centralized Rack History tab
 * Keeping constants for backward compatibility and migration
 */
var META_STATUS_COL = 6;        // Column F - Sync status (DEPRECATED - use History tab)
var META_ARENA_GUID_COL = 7;    // Column G - Arena item GUID (DEPRECATED - use History tab)
var META_LAST_SYNC_COL = 8;     // Column H - Last sync timestamp (DEPRECATED - use History tab)
var META_CHECKSUM_COL = 9;      // Column I - BOM checksum (DEPRECATED - use History tab)

// ============================================================================
// CORE STATUS MANAGEMENT FUNCTIONS
// ============================================================================

/**
 * Updates the sync status of a rack configuration sheet
 * REFACTORED: Now uses History tab instead of Row 1 metadata
 * @param {Sheet} sheet - Rack configuration sheet
 * @param {string} status - Status value from RACK_STATUS
 * @param {string} arenaGuid - Arena item GUID (optional, null for PLACEHOLDER)
 * @param {Object} eventDetails - Optional details for history event logging
 */
function updateRackSheetStatus(sheet, status, arenaGuid, eventDetails) {
  if (!sheet) {
    Logger.log('updateRackSheetStatus: No sheet provided');
    return;
  }

  // Verify it's a rack config sheet
  if (!isRackConfigSheet(sheet)) {
    Logger.log('updateRackSheetStatus: Not a rack config sheet: ' + sheet.getName());
    return;
  }

  try {
    var metadata = getRackConfigMetadata(sheet);
    var itemNumber = metadata.itemNumber;
    var rackName = metadata.itemName;

    // Get previous status for history logging
    var previousStatus = getRackStatusFromHistory(itemNumber);

    // Prepare metadata update
    var metadataUpdate = {
      status: status,
      lastSync: new Date()
    };

    if (arenaGuid !== undefined) {
      metadataUpdate.arenaGuid = arenaGuid;
    }

    // Calculate and store checksum if status is SYNCED
    if (status === RACK_STATUS.SYNCED) {
      var checksum = calculateBOMChecksum(sheet);
      metadataUpdate.checksum = checksum;
    }

    // Update History tab
    updateRackHistorySummary(itemNumber, rackName, metadataUpdate);

    // Log status change event to history
    if (previousStatus !== status) {
      var historyDetails = eventDetails || {};
      historyDetails.statusBefore = previousStatus || '';
      historyDetails.statusAfter = status;

      addRackHistoryEvent(itemNumber, HISTORY_EVENT.STATUS_CHANGE, historyDetails);
    }

    // Update tab name with status indicator
    updateRackTabName(sheet);

    Logger.log('✓ Updated rack status: ' + sheet.getName() + ' → ' + status);

  } catch (error) {
    Logger.log('Error updating rack status: ' + error.message);
  }
}

/**
 * Gets the current sync status of a rack configuration sheet
 * REFACTORED: Now reads from History tab instead of Row 1
 * @param {Sheet} sheet - Rack configuration sheet
 * @return {string|null} Status value or null if not found
 */
function getRackSheetStatus(sheet) {
  if (!sheet) return null;

  try {
    var metadata = getRackConfigMetadata(sheet);
    if (!metadata) return null;

    return getRackStatusFromHistory(metadata.itemNumber);
  } catch (error) {
    Logger.log('Error reading rack status: ' + error.message);
    return null;
  }
}

/**
 * Updates the tab name of a rack sheet with status indicator
 * Tries emoji first, falls back to text if emoji fails
 * @param {Sheet} sheet - Rack configuration sheet
 * @return {boolean} True if successful
 */
function updateRackTabName(sheet) {
  var metadata = getRackConfigMetadata(sheet);
  if (!metadata) {
    Logger.log('updateRackTabName: No metadata found for sheet');
    return false;
  }

  var status = getRackSheetStatus(sheet);
  if (!status) {
    Logger.log('updateRackTabName: No status found for ' + metadata.itemNumber);
    return false;
  }

  var currentName = sheet.getName();
  var indicator = STATUS_INDICATORS[status] || '';
  var baseName = 'Rack - ' + metadata.itemNumber + ' (' + metadata.itemName + ')';
  var newName = indicator + ' ' + baseName;

  Logger.log('updateRackTabName: ' + metadata.itemNumber);
  Logger.log('  Current name: "' + currentName + '"');
  Logger.log('  Status: ' + status);
  Logger.log('  Indicator emoji: "' + indicator + '"');
  Logger.log('  New name: "' + newName + '"');

  // Google Sheets tab name limit: 100 characters
  if (newName.length > 100) {
    newName = newName.substring(0, 97) + '...';
  }

  try {
    sheet.setName(newName);
    var actualName = sheet.getName();
    Logger.log('  ✓ Name set successfully to: "' + actualName + '"');
    return true;
  } catch (e) {
    // Emoji might not be supported, try text fallback
    Logger.log('  ⚠ Emoji failed in tab name: ' + e.message);
    Logger.log('  Trying text indicator fallback...');

    var textIndicator = STATUS_TEXT_INDICATORS[status] || '[???]';
    newName = textIndicator + ' ' + baseName;

    if (newName.length > 100) {
      newName = newName.substring(0, 97) + '...';
    }

    try {
      sheet.setName(newName);
      var actualName = sheet.getName();
      Logger.log('  ✓ Fallback name set to: "' + actualName + '"');
      return true;
    } catch (e2) {
      Logger.log('  ✗ Tab name update failed completely: ' + e2.message);
      return false;
    }
  }
}

/**
 * Calculates a checksum of the rack BOM data for change detection
 * Uses simple concatenation of key BOM fields
 * @param {Sheet} sheet - Rack configuration sheet
 * @return {string} Checksum string
 */
function calculateBOMChecksum(sheet) {
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) {
      return ''; // No BOM data
    }

    var data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 6).getValues();

    // Build checksum from item numbers and quantities
    var checksumParts = [];
    data.forEach(function(row) {
      var itemNumber = row[0]; // Column A
      var quantity = row[5];    // Column F (Qty)

      if (itemNumber && itemNumber.toString().trim() !== '') {
        // Format: "ITEM-NUMBER:QTY"
        checksumParts.push(itemNumber + ':' + (quantity || 1));
      }
    });

    return checksumParts.join('|');

  } catch (error) {
    Logger.log('Error calculating BOM checksum: ' + error.message);
    return '';
  }
}

/**
 * Detects if the rack sheet has been modified locally since last sync
 * REFACTORED: Compares current BOM checksum with stored checksum from History tab
 * @param {Sheet} sheet - Rack configuration sheet
 * @return {boolean} True if modified locally
 */
function detectLocalChanges(sheet) {
  var metadata = getRackConfigMetadata(sheet);
  if (!metadata) return false;

  var currentChecksum = calculateBOMChecksum(sheet);
  var storedChecksum = getRackChecksumFromHistory(metadata.itemNumber);

  if (!storedChecksum) {
    return false; // No baseline to compare against
  }

  return currentChecksum !== storedChecksum;
}

// ============================================================================
// OVERVIEW SHEET STATUS MANAGEMENT
// ============================================================================

/**
 * Updates the sync status of an overview sheet
 * @param {Sheet} sheet - Overview sheet
 * @param {string} status - Status value
 * @param {string} podGuid - POD item GUID (optional)
 */
function updateOverviewSheetStatus(sheet, status, podGuid) {
  if (!sheet) return;

  try {
    // Check if metadata row exists (Row 1 might not be set up for overview)
    var metaLabel = sheet.getRange(1, 1).getValue();
    if (metaLabel !== 'OVERVIEW_METADATA') {
      // Initialize overview metadata
      sheet.getRange(1, 1).setValue('OVERVIEW_METADATA');
      sheet.getRange(1, 1, 1, 10).setBackground('#f3f3f3').setFontWeight('bold');
    }

    // Store status in metadata
    sheet.getRange(1, 6).setValue(status); // Column F
    if (podGuid) {
      sheet.getRange(1, 7).setValue(podGuid); // Column G
    }
    sheet.getRange(1, 8).setValue(new Date()); // Column H - timestamp

    // Update tab name
    updateOverviewTabName(sheet, status);

    Logger.log('✓ Updated overview status: ' + sheet.getName() + ' → ' + status);

  } catch (error) {
    Logger.log('Error updating overview status: ' + error.message);
  }
}

/**
 * Gets the current status of an overview sheet
 * @param {Sheet} sheet - Overview sheet
 * @return {string|null} Status value or null
 */
function getOverviewSheetStatus(sheet) {
  if (!sheet) return null;

  try {
    var metaLabel = sheet.getRange(1, 1).getValue();
    if (metaLabel !== 'OVERVIEW_METADATA') {
      return null; // Not initialized
    }

    return sheet.getRange(1, 6).getValue() || null;
  } catch (error) {
    return null;
  }
}

/**
 * Updates the tab name of an overview sheet with status indicator
 * @param {Sheet} sheet - Overview sheet
 * @param {string} status - Status value
 */
function updateOverviewTabName(sheet, status) {
  var sheetName = sheet.getName();

  // Remove existing indicator if present
  var baseName = sheetName.replace(/^(\u{1F534}|\u{1F7E2}|\u{1F7E1}|\[NEW\]|\[OK\]|\[DIFF\])\s+/, '');

  var indicator = STATUS_INDICATORS[status] || '';
  var newName = indicator + ' ' + baseName;

  if (newName.length > 100) {
    newName = newName.substring(0, 97) + '...';
  }

  try {
    sheet.setName(newName);
  } catch (e) {
    // Try text fallback
    var textIndicator = STATUS_TEXT_INDICATORS[status] || '';
    newName = textIndicator + ' ' + baseName;
    try {
      sheet.setName(newName);
    } catch (e2) {
      Logger.log('Failed to update overview tab name: ' + e2.message);
    }
  }
}

// ============================================================================
// BATCH STATUS CHECKING
// ============================================================================

/**
 * Checks the status of all rack configuration sheets against Arena
 * Uses batch API calls with caching for performance
 * @return {Object} Summary of status check results
 */
function checkAllRackStatuses() {
  Logger.log('=== CHECK ALL RACK STATUSES START ===');

  var client = new ArenaAPIClient();
  var results = {
    synced: 0,
    outOfSync: 0,
    localModified: 0,
    placeholder: 0,
    error: 0,
    total: 0
  };

  try {
    // Get all rack configuration sheets
    var racks = getAllRackConfigTabs();
    results.total = racks.length;

    Logger.log('Found ' + racks.length + ' rack configuration sheets');

    if (racks.length === 0) {
      SpreadsheetApp.getUi().alert('No Racks Found', 'No rack configuration sheets found in this spreadsheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return results;
    }

    // Build map of Arena GUIDs to check
    var guidsToCheck = [];
    var rackByGuid = {};

    racks.forEach(function(rack) {
      var status = getRackSheetStatus(rack.sheet);

      Logger.log('Checking rack: ' + rack.itemNumber + ' with status: "' + status + '"');

      // Handle legacy racks (created before status feature)
      if (!status || status === '') {
        Logger.log('Legacy rack detected (no status metadata): ' + rack.itemNumber);
        // Try to look up in Arena by item number
        try {
          var item = client.getItemByNumber(rack.itemNumber);
          if (item && (item.guid || item.Guid)) {
            var guid = item.guid || item.Guid;
            Logger.log('  Found in Arena with GUID: ' + guid);
            // Initialize status metadata for this legacy rack
            updateRackSheetStatus(rack.sheet, RACK_STATUS.SYNCED, guid);
            guidsToCheck.push(guid);
            rackByGuid[guid] = rack;
          } else {
            Logger.log('  Not found in Arena - treating as placeholder');
            updateRackSheetStatus(rack.sheet, RACK_STATUS.PLACEHOLDER, null);
            results.placeholder++;
          }
        } catch (error) {
          Logger.log('  Error looking up legacy rack: ' + error.message);
          if (error.message && error.message.indexOf('404') !== -1) {
            // 404 = not found, it's a placeholder
            updateRackSheetStatus(rack.sheet, RACK_STATUS.PLACEHOLDER, null);
            results.placeholder++;
          } else {
            // Other error, skip this rack
            results.error++;
          }
        }
        return;
      }

      if (status === RACK_STATUS.PLACEHOLDER) {
        Logger.log('  → Placeholder rack, skipping Arena check');
        // Still update tab name to show red dot indicator
        updateRackTabName(rack.sheet);
        results.placeholder++;
        return; // Skip Arena API check for placeholders
      }

      // REFACTORED: Read GUID from History tab
      var arenaGuid = getRackArenaGuidFromHistory(rack.itemNumber);
      if (arenaGuid) {
        Logger.log('  → Has Arena GUID, will check against Arena');
        guidsToCheck.push(arenaGuid);
        rackByGuid[arenaGuid] = rack;
      } else {
        // Has status but no GUID - this is an error state
        Logger.log('⚠ Rack has status "' + status + '" but no Arena GUID: ' + rack.itemNumber);
        Logger.log('  This rack should have status PLACEHOLDER if not yet in Arena');
        results.error++;
      }
    });

    Logger.log('Checking ' + guidsToCheck.length + ' racks against Arena...');

    // Check each rack against Arena
    guidsToCheck.forEach(function(guid) {
      var rack = rackByGuid[guid];

      try {
        // Fetch Arena BOM
        var arenaBOM = client.makeRequest('/items/' + guid + '/bom', { method: 'GET' });
        var arenaBOMLines = arenaBOM.results || arenaBOM.Results || [];

        // Get current sheet BOM
        var sheetBOM = getCurrentRackBOMData(rack.sheet);

        // Compare (pass client to fetch full item details)
        var changes = compareBOMs(sheetBOM, arenaBOMLines, client);
        var hasChanges = changes.modified.length > 0 || changes.added.length > 0 || changes.removed.length > 0;

        if (hasChanges) {
          // Determine if local or Arena modified
          if (detectLocalChanges(rack.sheet)) {
            updateRackSheetStatus(rack.sheet, RACK_STATUS.LOCAL_MODIFIED, guid);
            results.localModified++;
          } else {
            updateRackSheetStatus(rack.sheet, RACK_STATUS.ARENA_MODIFIED, guid);
            results.outOfSync++;
          }
        } else {
          updateRackSheetStatus(rack.sheet, RACK_STATUS.SYNCED, guid);
          results.synced++;
        }

      } catch (error) {
        Logger.log('Error checking rack ' + rack.itemNumber + ': ' + error.message);
        updateRackSheetStatus(rack.sheet, RACK_STATUS.ERROR, guid);
        results.error++;
      }
    });

    // Show summary dialog
    showStatusCheckSummary(results);

  } catch (error) {
    Logger.log('Error in checkAllRackStatuses: ' + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Failed to check rack statuses: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }

  Logger.log('=== CHECK ALL RACK STATUSES COMPLETE ===');
  return results;
}

/**
 * Shows a summary dialog of status check results
 * @param {Object} results - Results object from checkAllRackStatuses
 */
function showStatusCheckSummary(results) {
  var ui = SpreadsheetApp.getUi();

  var message = 'Status check complete for ' + results.total + ' rack sheets:\n\n';
  message += '🟢 Synced: ' + results.synced + '\n';
  message += '🟡 Arena Modified: ' + results.outOfSync + '\n';
  message += '🟠 Locally Modified: ' + results.localModified + '\n';
  message += '🔴 Placeholder: ' + results.placeholder + '\n';

  if (results.error > 0) {
    message += '❌ Errors: ' + results.error + '\n';
  }

  message += '\n';

  if (results.outOfSync > 0) {
    message += 'TIP: Use "Refresh BOM" on yellow racks to see Arena changes.';
  }

  if (results.localModified > 0) {
    message += 'NOTE: Orange racks have local edits not yet pushed to Arena.';
  }

  ui.alert('Rack Status Check Results', message, ui.ButtonSet.OK);
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Gets current BOM data from a rack sheet in standardized format
 * @param {Sheet} sheet - Rack configuration sheet
 * @return {Array} Array of BOM line objects
 */
function getCurrentRackBOMData(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    return [];
  }

  var data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 6).getValues();
  var bomLines = [];

  data.forEach(function(row, index) {
    var itemNumber = row[0];
    if (!itemNumber || itemNumber.toString().trim() === '') {
      return; // Skip empty rows
    }

    bomLines.push({
      rowNumber: DATA_START_ROW + index, // Actual sheet row number
      itemNumber: itemNumber,
      name: row[1] || '',
      description: row[2] || '',
      category: row[3] || '',
      lifecycle: row[4] || '',
      quantity: row[5] || 1
    });
  });

  return bomLines;
}

/**
 * Manually mark a rack as synced (user override)
 * REFACTORED: Reads GUID from History tab and logs manual sync event
 * Useful if user knows status is incorrect
 */
function markCurrentRackAsSynced() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (!isRackConfigSheet(sheet)) {
    SpreadsheetApp.getUi().alert('Error', 'Current sheet is not a rack configuration sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var metadata = getRackConfigMetadata(sheet);
  if (!metadata) {
    SpreadsheetApp.getUi().alert('Error', 'Could not read rack metadata.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // REFACTORED: Read GUID from History tab
  var arenaGuid = getRackArenaGuidFromHistory(metadata.itemNumber);
  if (!arenaGuid) {
    SpreadsheetApp.getUi().alert('Error', 'This rack has no Arena GUID. Cannot mark as synced.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Update status with event details
  var eventDetails = {
    changesSummary: 'User manually marked as synced',
    details: 'User override - marked as synced via menu action'
  };

  updateRackSheetStatus(sheet, RACK_STATUS.SYNCED, arenaGuid, eventDetails);

  // Log manual sync event
  addRackHistoryEvent(metadata.itemNumber, HISTORY_EVENT.MANUAL_SYNC, {
    details: 'User manually marked rack as synced',
    statusAfter: RACK_STATUS.SYNCED
  });

  SpreadsheetApp.getUi().alert('Success', 'Rack marked as synced with Arena.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Refreshes all rack tab names based on current status in History tab
 * Useful for fixing missing status indicators after migration or updates
 */
function refreshAllRackTabNames() {
  try {
    var racks = getAllRackConfigTabs();
    var updated = 0;
    var skipped = 0;

    Logger.log('Refreshing tab names for ' + racks.length + ' racks');

    racks.forEach(function(rack) {
      var success = updateRackTabName(rack.sheet);
      if (success) {
        updated++;
        Logger.log('  ✓ Updated: ' + rack.itemNumber);
      } else {
        skipped++;
        Logger.log('  ○ Skipped: ' + rack.itemNumber + ' (no status)');
      }
    });

    var message = 'Rack tab names refreshed!\n\n';
    message += 'Updated: ' + updated + '\n';
    message += 'Skipped: ' + skipped + ' (no status set)\n';

    SpreadsheetApp.getUi().alert('Tab Names Refreshed', message, SpreadsheetApp.getUi().ButtonSet.OK);

    Logger.log('Tab name refresh complete: ' + updated + ' updated, ' + skipped + ' skipped');

  } catch (error) {
    Logger.log('Error refreshing rack tab names: ' + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Failed to refresh tab names: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
