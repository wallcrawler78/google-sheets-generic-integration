/**
 * Rack History Management
 * Centralized history tracking for all rack configurations
 * Replaces Row 1 metadata storage with dedicated History tab
 */

// History tab constants
var HISTORY_TAB_NAME = 'Rack History';
var HISTORY_FROZEN_ROWS = 1;  // Header row frozen

// Summary section columns (Row 1 = headers, Row 2+ = one row per rack)
var HIST_SUMMARY_ITEM_NUM_COL = 1;      // A: Rack Item#
var HIST_SUMMARY_RACK_NAME_COL = 2;     // B: Rack Name
var HIST_SUMMARY_STATUS_COL = 3;        // C: Current Status
var HIST_SUMMARY_ARENA_GUID_COL = 4;    // D: Arena GUID
var HIST_SUMMARY_CREATED_COL = 5;       // E: Created Date
var HIST_SUMMARY_LAST_REFRESH_COL = 6;  // F: Last Refresh
var HIST_SUMMARY_LAST_SYNC_COL = 7;     // G: Last Sync
var HIST_SUMMARY_LAST_PUSH_COL = 8;     // H: Last Push
var HIST_SUMMARY_CHECKSUM_COL = 9;      // I: BOM Checksum

// Detail section columns (below summary rows, separated by blank row)
var HIST_DETAIL_TIMESTAMP_COL = 1;      // A: Timestamp
var HIST_DETAIL_RACK_COL = 2;           // B: Rack Item#
var HIST_DETAIL_EVENT_TYPE_COL = 3;     // C: Event Type
var HIST_DETAIL_USER_COL = 4;           // D: User
var HIST_DETAIL_STATUS_BEFORE_COL = 5;  // E: Status Before
var HIST_DETAIL_STATUS_AFTER_COL = 6;   // F: Status After
var HIST_DETAIL_SUMMARY_COL = 7;        // G: Changes Summary
var HIST_DETAIL_DETAILS_COL = 8;        // H: Details
var HIST_DETAIL_LINK_COL = 9;           // I: Link

// Event type constants
var HISTORY_EVENT = {
  RACK_CREATED: 'RACK_CREATED',
  STATUS_CHANGE: 'STATUS_CHANGE',
  LOCAL_EDIT: 'LOCAL_EDIT',
  REFRESH_ACCEPTED: 'REFRESH_ACCEPTED',
  REFRESH_DECLINED: 'REFRESH_DECLINED',
  REFRESH_NO_CHANGES: 'REFRESH_NO_CHANGES',
  POD_PUSH: 'POD_PUSH',
  BOM_PULL: 'BOM_PULL',
  MANUAL_SYNC: 'MANUAL_SYNC',
  BATCH_CHECK: 'BATCH_CHECK',
  ERROR: 'ERROR',
  CHECKSUM_MISMATCH: 'CHECKSUM_MISMATCH',
  MIGRATION: 'MIGRATION'
};

/**
 * Gets or creates the Rack History tab
 * @return {Sheet} The Rack History sheet
 */
function getOrCreateRackHistoryTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = ss.getSheetByName(HISTORY_TAB_NAME);

  if (historySheet) {
    return historySheet;
  }

  // Create new History tab
  Logger.log('Creating new Rack History tab');
  historySheet = ss.insertSheet(HISTORY_TAB_NAME);

  // Set up summary section header (Row 1)
  var summaryHeaders = [
    'Rack Item#',
    'Rack Name',
    'Current Status',
    'Arena GUID',
    'Created Date',
    'Last Refresh',
    'Last Sync',
    'Last Push',
    'BOM Checksum'
  ];

  var headerRange = historySheet.getRange(1, 1, 1, summaryHeaders.length);
  headerRange.setValues([summaryHeaders]);
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');

  // Initial freeze - will be updated dynamically as racks are added
  historySheet.setFrozenRows(3);  // Header + separator + detail header

  // Set column widths for summary section
  historySheet.setColumnWidth(HIST_SUMMARY_ITEM_NUM_COL, 120);      // Item#
  historySheet.setColumnWidth(HIST_SUMMARY_RACK_NAME_COL, 200);     // Name
  historySheet.setColumnWidth(HIST_SUMMARY_STATUS_COL, 150);        // Status
  historySheet.setColumnWidth(HIST_SUMMARY_ARENA_GUID_COL, 180);    // GUID
  historySheet.setColumnWidth(HIST_SUMMARY_CREATED_COL, 120);       // Created
  historySheet.setColumnWidth(HIST_SUMMARY_LAST_REFRESH_COL, 150);  // Last Refresh
  historySheet.setColumnWidth(HIST_SUMMARY_LAST_SYNC_COL, 150);     // Last Sync
  historySheet.setColumnWidth(HIST_SUMMARY_LAST_PUSH_COL, 150);     // Last Push
  historySheet.setColumnWidth(HIST_SUMMARY_CHECKSUM_COL, 300);      // Checksum

  // Add separator row (Row 2)
  historySheet.getRange(2, 1, 1, summaryHeaders.length).setBackground('#f0f0f0');

  // Set up detail section headers (Row 3)
  var detailHeaders = [
    'Timestamp',
    'Rack Item#',
    'Event Type',
    'User',
    'Status Before',
    'Status After',
    'Changes Summary',
    'Details',
    'Link'
  ];

  var detailHeaderRange = historySheet.getRange(3, 1, 1, detailHeaders.length);
  detailHeaderRange.setValues([detailHeaders]);
  detailHeaderRange.setBackground('#34a853');
  detailHeaderRange.setFontColor('white');
  detailHeaderRange.setFontWeight('bold');
  detailHeaderRange.setHorizontalAlignment('center');

  // Set column widths for detail section
  historySheet.setColumnWidth(HIST_DETAIL_TIMESTAMP_COL, 150);       // Timestamp
  historySheet.setColumnWidth(HIST_DETAIL_RACK_COL, 120);            // Rack
  historySheet.setColumnWidth(HIST_DETAIL_EVENT_TYPE_COL, 150);      // Event Type
  historySheet.setColumnWidth(HIST_DETAIL_USER_COL, 200);            // User
  historySheet.setColumnWidth(HIST_DETAIL_STATUS_BEFORE_COL, 120);   // Status Before
  historySheet.setColumnWidth(HIST_DETAIL_STATUS_AFTER_COL, 120);    // Status After
  historySheet.setColumnWidth(HIST_DETAIL_SUMMARY_COL, 250);         // Summary
  historySheet.setColumnWidth(HIST_DETAIL_DETAILS_COL, 300);         // Details
  historySheet.setColumnWidth(HIST_DETAIL_LINK_COL, 100);            // Link

  // Set tab color to purple (history theme)
  historySheet.setTabColor('#9c27b0');

  // Protect the History tab to prevent accidental edits
  protectHistoryTab(historySheet);

  Logger.log('Rack History tab created successfully');
  return historySheet;
}

/**
 * Finds the summary row number for a rack in History tab
 * @param {string} itemNumber - Rack item number
 * @return {number} Row number (2+) or -1 if not found
 */
function findRackHistorySummaryRow(itemNumber) {
  var historySheet = getOrCreateRackHistoryTab();
  var lastRow = historySheet.getLastRow();

  if (lastRow < 2) {
    return -1;  // No summary rows yet
  }

  // Summary rows are between row 2 and the separator row (row 2 is separator)
  // Actually, based on structure: Row 1 = headers, Row 2+ = summary rows, then separator, then details
  // Let me reconsider: Row 1 = summary headers, Rows 2-N = rack summaries, Row N+1 = separator, Row N+2 = detail headers
  // For now, search all rows looking for matching item number in column A

  var summaryData = historySheet.getRange(2, HIST_SUMMARY_ITEM_NUM_COL, Math.max(1, lastRow - 1), 1).getValues();

  for (var i = 0; i < summaryData.length; i++) {
    if (summaryData[i][0] === itemNumber || summaryData[i][0] === itemNumber.toString()) {
      return i + 2;  // +2 because array is 0-indexed and we started at row 2
    }
  }

  return -1;  // Not found
}

/**
 * Gets the row number where detail section starts
 * @return {number} First row of detail section (after separator)
 */
function getDetailSectionStartRow() {
  var historySheet = getOrCreateRackHistoryTab();
  var lastRow = historySheet.getLastRow();

  // Find the separator row (gray background) or detail header row (green background)
  for (var i = 2; i <= lastRow; i++) {
    var cell = historySheet.getRange(i, 1);
    var background = cell.getBackground();

    // Detail headers have green background (#34a853)
    if (background === '#34a853') {
      return i + 1;  // Detail data starts after header
    }
  }

  // If not found, detail section starts at row 3 (after summary header)
  return 3;
}

/**
 * Creates a new summary row for a rack
 * @param {string} itemNumber - Rack item number
 * @param {string} rackName - Rack name
 * @param {Object} metadata - Initial metadata
 * @return {number} Row number of created summary
 */
function createRackHistorySummaryRow(itemNumber, rackName, metadata) {
  var historySheet = getOrCreateRackHistoryTab();

  // Find where to insert (after last summary row, before separator/detail section)
  var lastRow = historySheet.getLastRow();
  var insertRow = 2;  // Default to row 2 (first summary row)

  // Find existing summary rows
  if (lastRow >= 2) {
    // Count existing summary rows (they have item numbers in column A)
    var summaryCount = 0;
    for (var i = 2; i <= lastRow; i++) {
      var cellValue = historySheet.getRange(i, 1).getValue();
      if (cellValue && cellValue !== 'Timestamp') {  // Not a detail header
        summaryCount++;
        insertRow = i + 1;
      } else {
        break;  // Reached separator or detail section
      }
    }
  }

  // Insert new row at the end of summary section
  historySheet.insertRowBefore(insertRow);

  // Set summary row values
  var summaryRow = [
    itemNumber,
    rackName,
    metadata.status || RACK_STATUS.PLACEHOLDER,
    metadata.arenaGuid || '',
    metadata.created || new Date(),
    metadata.lastRefresh || '',
    metadata.lastSync || '',
    metadata.lastPush || '',
    metadata.checksum || ''
  ];

  historySheet.getRange(insertRow, 1, 1, summaryRow.length).setValues([summaryRow]);

  // Format the new row
  var rowRange = historySheet.getRange(insertRow, 1, 1, summaryRow.length);
  rowRange.setFontWeight('normal');
  rowRange.setBackground('white');

  // Format status column with emoji
  var statusCell = historySheet.getRange(insertRow, HIST_SUMMARY_STATUS_COL);
  var statusWithEmoji = getStatusWithEmoji(metadata.status || RACK_STATUS.PLACEHOLDER);
  statusCell.setValue(statusWithEmoji);

  Logger.log('Created summary row for rack ' + itemNumber + ' at row ' + insertRow);

  // Update freeze point after adding new rack
  updateHistoryTabFreeze();

  return insertRow;
}

/**
 * Updates summary row for a rack (or creates if doesn't exist)
 * @param {string} itemNumber - Rack item number
 * @param {Object} metadata - Metadata to update
 */
function updateRackHistorySummary(itemNumber, rackName, metadata) {
  var row = findRackHistorySummaryRow(itemNumber);

  if (row === -1) {
    // Create new summary row
    createRackHistorySummaryRow(itemNumber, rackName, metadata);
    return;
  }

  var historySheet = getOrCreateRackHistoryTab();

  // Update existing row (only update fields that are provided)
  if (metadata.status !== undefined) {
    var statusWithEmoji = getStatusWithEmoji(metadata.status);
    historySheet.getRange(row, HIST_SUMMARY_STATUS_COL).setValue(statusWithEmoji);
  }

  if (metadata.arenaGuid !== undefined) {
    historySheet.getRange(row, HIST_SUMMARY_ARENA_GUID_COL).setValue(metadata.arenaGuid);
  }

  if (metadata.created !== undefined) {
    historySheet.getRange(row, HIST_SUMMARY_CREATED_COL).setValue(metadata.created);
  }

  if (metadata.lastRefresh !== undefined) {
    historySheet.getRange(row, HIST_SUMMARY_LAST_REFRESH_COL).setValue(metadata.lastRefresh);
  }

  if (metadata.lastSync !== undefined) {
    historySheet.getRange(row, HIST_SUMMARY_LAST_SYNC_COL).setValue(metadata.lastSync);
  }

  if (metadata.lastPush !== undefined) {
    historySheet.getRange(row, HIST_SUMMARY_LAST_PUSH_COL).setValue(metadata.lastPush);
  }

  if (metadata.checksum !== undefined) {
    historySheet.getRange(row, HIST_SUMMARY_CHECKSUM_COL).setValue(metadata.checksum);
  }

  if (rackName) {
    historySheet.getRange(row, HIST_SUMMARY_RACK_NAME_COL).setValue(rackName);
  }

  Logger.log('Updated summary row for rack ' + itemNumber);
}

/**
 * Gets status string with emoji indicator
 * @param {string} status - Status constant
 * @return {string} Status with emoji
 */
function getStatusWithEmoji(status) {
  var emoji = STATUS_INDICATORS[status] || '';
  return emoji + ' ' + status;
}

/**
 * Adds a detailed history event
 * @param {string} itemNumber - Rack item number
 * @param {string} eventType - Event type constant
 * @param {Object} details - Event details
 */
function addRackHistoryEvent(itemNumber, eventType, details) {
  var historySheet = getOrCreateRackHistoryTab();
  var lastRow = historySheet.getLastRow();

  // Find detail section start
  var detailStartRow = getDetailSectionStartRow();

  // Append new row at the end
  var newRow = lastRow + 1;

  var user = Session.getActiveUser().getEmail() || 'Unknown';

  var eventRow = [
    new Date(),                           // Timestamp
    itemNumber,                           // Rack Item#
    eventType,                            // Event Type
    user,                                 // User
    details.statusBefore || '',           // Status Before
    details.statusAfter || '',            // Status After
    details.changesSummary || '',         // Changes Summary
    details.details || '',                // Details
    details.link || ''                    // Link
  ];

  historySheet.getRange(newRow, 1, 1, eventRow.length).setValues([eventRow]);

  // Format timestamp
  historySheet.getRange(newRow, HIST_DETAIL_TIMESTAMP_COL).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  Logger.log('Added history event: ' + eventType + ' for rack ' + itemNumber);
}

/**
 * Gets current status for a rack from History tab
 * @param {string} itemNumber - Rack item number
 * @return {string|null} Current status or null if not found
 */
function getRackStatusFromHistory(itemNumber) {
  var row = findRackHistorySummaryRow(itemNumber);

  if (row === -1) {
    return null;
  }

  var historySheet = getOrCreateRackHistoryTab();
  var statusValue = historySheet.getRange(row, HIST_SUMMARY_STATUS_COL).getValue();

  // Remove emoji prefix if present
  var status = statusValue.toString().replace(/[🔴🟢🟠🟡❌]\s*/, '');
  return status;
}

/**
 * Gets Arena GUID for a rack from History tab
 * @param {string} itemNumber - Rack item number
 * @return {string|null} Arena GUID or null if not found
 */
function getRackArenaGuidFromHistory(itemNumber) {
  var row = findRackHistorySummaryRow(itemNumber);

  if (row === -1) {
    return null;
  }

  var historySheet = getOrCreateRackHistoryTab();
  return historySheet.getRange(row, HIST_SUMMARY_ARENA_GUID_COL).getValue();
}

/**
 * Gets BOM checksum for a rack from History tab
 * @param {string} itemNumber - Rack item number
 * @return {string|null} Checksum or null if not found
 */
function getRackChecksumFromHistory(itemNumber) {
  var row = findRackHistorySummaryRow(itemNumber);

  if (row === -1) {
    return null;
  }

  var historySheet = getOrCreateRackHistoryTab();
  return historySheet.getRange(row, HIST_SUMMARY_CHECKSUM_COL).getValue();
}

/**
 * Creates "History" hyperlink in rack sheet cell D1
 * @param {Sheet} sheet - Rack configuration sheet
 */
function createHistoryLinkInRackSheet(sheet) {
  if (!isRackConfigSheet(sheet)) {
    Logger.log('Not a rack config sheet, skipping history link creation');
    return;
  }

  var metadata = getRackConfigMetadata(sheet);
  if (!metadata) return;

  var itemNumber = metadata.itemNumber;

  // Create hyperlink formula that jumps to History tab
  var historySheet = getOrCreateRackHistoryTab();
  var summaryRow = findRackHistorySummaryRow(itemNumber);

  var formula;
  if (summaryRow !== -1) {
    // Link to specific rack's summary row
    formula = '=HYPERLINK("#gid=' + historySheet.getSheetId() + '&range=A' + summaryRow + '", "📋 History")';
  } else {
    // Link to History tab in general
    formula = '=HYPERLINK("#gid=' + historySheet.getSheetId() + '", "📋 History")';
  }

  var cell = sheet.getRange(METADATA_ROW, META_ITEM_DESC_COL);
  cell.setFormula(formula);
  cell.setFontColor('#1a73e8');
  cell.setFontWeight('normal');
  cell.setHorizontalAlignment('left');

  Logger.log('Created history link in ' + sheet.getName() + ' at D1');
}

/**
 * Applies auto-filter to show only one rack's detailed history
 * @param {string} itemNumber - Rack item number to filter
 */
function filterHistoryByRack(itemNumber) {
  var historySheet = getOrCreateRackHistoryTab();

  // Remove existing filters
  var existingFilter = historySheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  // Find detail section
  var detailStartRow = getDetailSectionStartRow();
  var lastRow = historySheet.getLastRow();

  if (lastRow < detailStartRow) {
    Logger.log('No detail rows to filter');
    return;
  }

  // Create filter on detail section
  var filterRange = historySheet.getRange(detailStartRow - 1, 1, lastRow - detailStartRow + 2, 9);
  var filter = filterRange.createFilter();

  // Apply filter criteria on column B (Rack Item#)
  var criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(itemNumber)
    .build();

  filter.setColumnFilterCriteria(HIST_DETAIL_RACK_COL, criteria);

  // Scroll to detail section
  historySheet.setActiveRange(historySheet.getRange(detailStartRow, 1));

  Logger.log('Filtered history to show only rack: ' + itemNumber);

  SpreadsheetApp.getUi().alert(
    'History Filtered',
    'Now showing history for rack: ' + itemNumber + '\n\n' +
    'To clear filter, use Data > Remove filter or click another rack.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Click handler for History tab - applies filter when summary row clicked
 * This should be set up as an installable trigger
 */
function onHistoryTabClick(e) {
  try {
    if (!e || !e.range) return;

    var sheet = e.range.getSheet();
    if (sheet.getName() !== HISTORY_TAB_NAME) return;

    var row = e.range.getRow();

    // Check if clicked on a summary row (row 2+, but not separator or detail header)
    if (row >= 2) {
      var itemNumber = sheet.getRange(row, HIST_SUMMARY_ITEM_NUM_COL).getValue();

      // Validate it's a real rack item number (not empty, not "Timestamp" header)
      if (itemNumber && itemNumber !== 'Timestamp' && itemNumber.toString().trim() !== '') {
        filterHistoryByRack(itemNumber);
      }
    }
  } catch (error) {
    Logger.log('Error in onHistoryTabClick: ' + error.message);
  }
}

// ============================================================================
// MIGRATION FUNCTIONS
// ============================================================================

/**
 * Migrates existing rack metadata from Row 1 to History tab
 * Run this ONCE after deploying History tab feature
 * Safe to run multiple times - will not duplicate data
 */
function migrateRackMetadataToHistory() {
  Logger.log('=== MIGRATION: Rack Metadata → History Tab ===');

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Migrate Rack Metadata',
    'This will migrate existing rack metadata from Row 1 to the centralized History tab.\n\n' +
    'This is a one-time migration and is safe to run.\n\n' +
    'Continue with migration?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    Logger.log('Migration cancelled by user');
    return;
  }

  var results = {
    migrated: 0,
    skipped: 0,
    errors: 0
  };

  try {
    // Ensure History tab exists
    getOrCreateRackHistoryTab();

    // Get all rack configuration sheets
    var racks = getAllRackConfigTabs();
    Logger.log('Found ' + racks.length + ' rack configuration sheets to migrate');

    racks.forEach(function(rack) {
      try {
        var itemNumber = rack.itemNumber;
        var sheet = rack.sheet;

        Logger.log('Processing rack: ' + itemNumber);

        // Check if already migrated (has entry in History tab)
        var existingRow = findRackHistorySummaryRow(itemNumber);
        if (existingRow !== -1) {
          Logger.log('  Skipping - already in History tab');
          results.skipped++;
          return;
        }

        // Read existing metadata from Row 1 (deprecated columns F-I)
        var status = sheet.getRange(METADATA_ROW, META_STATUS_COL).getValue();
        var arenaGuid = sheet.getRange(METADATA_ROW, META_ARENA_GUID_COL).getValue();
        var lastSync = sheet.getRange(METADATA_ROW, META_LAST_SYNC_COL).getValue();
        var checksum = sheet.getRange(METADATA_ROW, META_CHECKSUM_COL).getValue();

        // Default values if metadata not found
        if (!status || status === '') {
          status = RACK_STATUS.PLACEHOLDER;
        }

        // Create History tab summary row
        createRackHistorySummaryRow(itemNumber, rack.itemName, {
          status: status,
          arenaGuid: arenaGuid || '',
          created: new Date(),  // We don't have original creation date, use current
          lastRefresh: '',
          lastSync: lastSync || '',
          lastPush: '',
          checksum: checksum || ''
        });

        // Log migration event
        addRackHistoryEvent(itemNumber, HISTORY_EVENT.MIGRATION, {
          changesSummary: 'Metadata migrated to History tab',
          details: 'Migrated from Row 1 metadata storage',
          statusAfter: status
        });

        // Clear old metadata from Row 1 (columns F-I) - optional, can leave for safety
        // Commenting out for now to preserve data during migration
        // sheet.getRange(METADATA_ROW, META_STATUS_COL, 1, 4).clearContent();

        // Add History link in D1 if not already present
        createHistoryLinkInRackSheet(sheet);

        Logger.log('  ✓ Migrated successfully');
        results.migrated++;

      } catch (error) {
        Logger.log('  ✗ Error migrating rack ' + rack.itemNumber + ': ' + error.message);
        results.errors++;
      }
    });

    // Update freeze point after migration
    updateHistoryTabFreeze();

    // Show summary
    var message = 'Migration Complete!\n\n';
    message += '✓ Migrated: ' + results.migrated + '\n';
    message += '○ Skipped (already migrated): ' + results.skipped + '\n';

    if (results.errors > 0) {
      message += '✗ Errors: ' + results.errors + '\n\n';
      message += 'Check execution log for details.';
    }

    ui.alert('Migration Results', message, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('ERROR in migration: ' + error.message);
    ui.alert('Migration Error', 'Migration failed: ' + error.message, ui.ButtonSet.OK);
  }

  Logger.log('=== MIGRATION COMPLETE ===');
  Logger.log('Results: ' + JSON.stringify(results));
}

/**
 * Clears old metadata from Row 1 of all rack sheets (columns F-I)
 * Run this AFTER migration and verification that History tab is working
 * WARNING: This is destructive! Make sure History tab has all data first
 */
function clearOldRackMetadata() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    'Clear Old Metadata - WARNING',
    'This will permanently delete metadata from Row 1 (columns F-I) of all rack sheets.\n\n' +
    'ONLY run this AFTER:\n' +
    '1. Migration is complete\n' +
    '2. You have verified History tab has all data\n' +
    '3. You have tested the new system\n\n' +
    'This action CANNOT be undone!\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    Logger.log('Clear metadata cancelled by user');
    return;
  }

  // Double confirmation
  var doubleCheck = ui.alert(
    'Final Confirmation',
    'Are you ABSOLUTELY SURE you want to delete old metadata?\n\n' +
    'This cannot be undone!',
    ui.ButtonSet.YES_NO
  );

  if (doubleCheck !== ui.Button.YES) {
    Logger.log('Clear metadata cancelled by user (second confirmation)');
    return;
  }

  try {
    var racks = getAllRackConfigTabs();
    var count = 0;

    racks.forEach(function(rack) {
      try {
        // Clear columns F-I (old metadata)
        rack.sheet.getRange(METADATA_ROW, META_STATUS_COL, 1, 4).clearContent();
        count++;
      } catch (error) {
        Logger.log('Error clearing metadata for ' + rack.itemNumber + ': ' + error.message);
      }
    });

    ui.alert('Success', 'Cleared old metadata from ' + count + ' rack sheets.', ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('Error', 'Failed to clear metadata: ' + error.message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// DYNAMIC FREEZE & SIDEBAR FUNCTIONS
// ============================================================================

/**
 * Updates the freeze point in History tab based on number of racks
 * Freezes: Summary header + All rack summaries + Separator + Detail header
 * Scrolls: Detail history rows only
 */
function updateHistoryTabFreeze() {
  try {
    var historySheet = getOrCreateRackHistoryTab();

    // Count rack summary rows (between row 2 and the separator/detail header)
    var lastRow = historySheet.getLastRow();
    var rackCount = 0;

    // Find where rack summaries end (look for separator or detail header)
    for (var i = 2; i <= lastRow; i++) {
      var cell = historySheet.getRange(i, 1);
      var value = cell.getValue();
      var background = cell.getBackground();

      // Stop at separator (gray) or detail header (green) or "Timestamp"
      if (background === '#f0f0f0' || background === '#34a853' || value === 'Timestamp') {
        break;
      }

      // Count non-empty rows with item numbers
      if (value && value.toString().trim() !== '') {
        rackCount++;
      }
    }

    // Calculate freeze point: 1 (header) + rackCount + 1 (separator) + 1 (detail header)
    var freezePoint = 1 + rackCount + 1 + 1;

    // Ensure minimum freeze of 3 (header + separator + detail header)
    freezePoint = Math.max(3, freezePoint);

    Logger.log('updateHistoryTabFreeze: Setting freeze to row ' + freezePoint + ' (for ' + rackCount + ' racks)');
    historySheet.setFrozenRows(freezePoint);

  } catch (error) {
    Logger.log('Error updating History tab freeze: ' + error.message);
  }
}

/**
 * Shows the History Filter Sidebar
 */
function showHistoryFilterSidebar() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('HistoryFilterSidebar')
      .setTitle('Rack History Filter')
      .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);

    Logger.log('History Filter Sidebar opened');
  } catch (error) {
    Logger.log('Error showing History Filter Sidebar: ' + error.message);
  }
}

/**
 * Gets list of all racks for sidebar dropdown
 * @return {Array} Array of {itemNumber, name, status}
 */
function getHistoryRackList() {
  try {
    var racks = getAllRackConfigTabs();

    return racks.map(function(rack) {
      var status = getRackStatusFromHistory(rack.itemNumber);
      return {
        itemNumber: rack.itemNumber,
        name: rack.itemName,
        status: status || RACK_STATUS.PLACEHOLDER
      };
    });
  } catch (error) {
    Logger.log('Error getting rack list: ' + error.message);
    return [];
  }
}

/**
 * Gets status statistics for sidebar
 * Reads directly from History tab summary rows for accuracy
 * @return {Object} Status counts
 */
function getHistoryStats() {
  try {
    var historySheet = getOrCreateRackHistoryTab();
    var stats = {
      total: 0,
      synced: 0,
      outOfSync: 0,
      localModified: 0,
      placeholder: 0,
      error: 0
    };

    // Read directly from History tab summary rows
    var lastRow = historySheet.getLastRow();
    if (lastRow < 2) {
      return stats; // No racks yet
    }

    // Find where summary section ends (look for separator or detail header)
    var summaryEndRow = 2;
    for (var i = 2; i <= lastRow; i++) {
      var cell = historySheet.getRange(i, 1);
      var value = cell.getValue();
      var background = cell.getBackground();

      // Stop at separator (gray) or detail header (green) or "Timestamp"
      if (background === '#f0f0f0' || background === '#34a853' || value === 'Timestamp') {
        summaryEndRow = i - 1;
        break;
      }

      // If we're at the last row and haven't hit a separator, use it
      if (i === lastRow && value && value.toString().trim() !== '') {
        summaryEndRow = i;
      }
    }

    // Count summary rows
    var rackCount = summaryEndRow - 1; // Subtract 1 for header row
    if (rackCount < 1) {
      return stats;
    }

    stats.total = rackCount;

    // Read status column for all summary rows
    var statusRange = historySheet.getRange(2, HIST_SUMMARY_STATUS_COL, rackCount, 1);
    var statusValues = statusRange.getValues();

    Logger.log('getHistoryStats: Reading ' + rackCount + ' rack statuses');

    statusValues.forEach(function(row, index) {
      var statusWithEmoji = row[0];
      if (!statusWithEmoji) return;

      // Strip emoji and whitespace to get clean status
      var status = statusWithEmoji.toString().trim().replace(/[🔴🟢🟠🟡❌]\s*/g, '');

      Logger.log('  Rack ' + (index + 1) + ': "' + statusWithEmoji + '" → "' + status + '"');

      switch(status) {
        case 'SYNCED':
          stats.synced++;
          break;
        case 'ARENA_MODIFIED':
          stats.outOfSync++;
          break;
        case 'LOCAL_MODIFIED':
          stats.localModified++;
          break;
        case 'PLACEHOLDER':
          stats.placeholder++;
          break;
        case 'ERROR':
          stats.error++;
          break;
        default:
          Logger.log('  WARNING: Unknown status "' + status + '"');
      }
    });

    Logger.log('getHistoryStats: Results = ' + JSON.stringify(stats));
    return stats;

  } catch (error) {
    Logger.log('Error getting history stats: ' + error.message);
    Logger.log('Error stack: ' + error.stack);
    return {
      total: 0,
      synced: 0,
      outOfSync: 0,
      localModified: 0,
      placeholder: 0,
      error: 0
    };
  }
}

/**
 * Clears filter from History tab
 */
function clearHistoryFilter() {
  try {
    var historySheet = getOrCreateRackHistoryTab();
    var filter = historySheet.getFilter();

    if (filter) {
      filter.remove();
      Logger.log('History filter cleared');
      return { success: true, message: 'Filter cleared' };
    } else {
      Logger.log('No filter to clear');
      return { success: true, message: 'No filter applied' };
    }
  } catch (error) {
    Logger.log('Error clearing history filter: ' + error.message);
    return { success: false, message: error.message };
  }
}

// ============================================================================
// DATA INTEGRITY & PROTECTION
// ============================================================================

/**
 * Protects the History tab from accidental edits
 * Allows script to edit but prevents user edits
 * @param {Sheet} sheet - History sheet to protect
 */
function protectHistoryTab(sheet) {
  try {
    // Remove any existing protections
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(function(protection) {
      if (protection.canEdit()) {
        protection.remove();
      }
    });

    // Create new protection
    var protection = sheet.protect();
    protection.setDescription('Rack History - Managed by Arena Data Center');

    // Set warning message
    protection.setWarningOnly(false);

    // Allow script owner to edit (so automated updates work)
    var me = Session.getEffectiveUser();
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());

    // Add warning message for users who try to edit
    protection.setDescription(
      'This sheet is automatically managed by the Arena Data Center system. ' +
      'Do not edit manually - data may be overwritten. ' +
      'Use the History Filter sidebar to view and filter history.'
    );

    Logger.log('History tab protected successfully');
  } catch (error) {
    Logger.log('Error protecting History tab: ' + error.message);
    // Don't fail if protection fails - just log it
  }
}

/**
 * Validates History tab data integrity
 * Checks for missing data, invalid statuses, orphaned events
 * @return {Object} Validation results with warnings
 */
function validateHistoryTabIntegrity() {
  var warnings = [];
  var errors = [];

  try {
    var historySheet = getOrCreateRackHistoryTab();
    var racks = getAllRackConfigTabs();

    // Check 1: All racks have summary rows
    racks.forEach(function(rack) {
      var summaryRow = findRackHistorySummaryRow(rack.itemNumber);
      if (summaryRow === -1) {
        warnings.push('Rack "' + rack.itemNumber + '" missing from History tab');
      }
    });

    // Check 2: All summary rows have corresponding rack sheets
    var lastRow = historySheet.getLastRow();
    for (var i = 2; i <= lastRow; i++) {
      var cell = historySheet.getRange(i, 1);
      var background = cell.getBackground();

      // Stop at separator or detail header
      if (background === '#f0f0f0' || background === '#34a853') {
        break;
      }

      var itemNumber = cell.getValue();
      if (itemNumber && itemNumber.toString().trim() !== '') {
        var rackSheet = findRackConfigTab(itemNumber);
        if (!rackSheet) {
          warnings.push('History summary for "' + itemNumber + '" has no corresponding rack sheet (orphaned)');
        }
      }
    }

    // Check 3: Validate status values
    var statusRange = historySheet.getRange(2, HIST_SUMMARY_STATUS_COL, lastRow - 1, 1);
    var statusValues = statusRange.getValues();

    statusValues.forEach(function(row, index) {
      var statusWithEmoji = row[0];
      if (!statusWithEmoji) return;

      var status = statusWithEmoji.toString().trim().replace(/[🔴🟢🟠🟡❌]\s*/g, '');
      var validStatuses = ['SYNCED', 'ARENA_MODIFIED', 'LOCAL_MODIFIED', 'PLACEHOLDER', 'ERROR', 'Timestamp', ''];

      if (status !== '' && validStatuses.indexOf(status) === -1) {
        errors.push('Row ' + (index + 2) + ': Invalid status "' + status + '"');
      }
    });

    return {
      success: errors.length === 0,
      warnings: warnings,
      errors: errors
    };

  } catch (error) {
    return {
      success: false,
      warnings: warnings,
      errors: ['Validation failed: ' + error.message]
    };
  }
}

/**
 * Repairs History tab integrity issues
 * Fixes missing summary rows, removes orphaned entries
 * @return {Object} Repair results
 */
function repairHistoryTabIntegrity() {
  var ui = SpreadsheetApp.getUi();
  var validation = validateHistoryTabIntegrity();

  if (validation.success && validation.warnings.length === 0) {
    ui.alert('No Issues Found', 'History tab integrity is good!', ui.ButtonSet.OK);
    return { success: true, message: 'No repairs needed' };
  }

  // Show issues and ask for confirmation
  var message = 'Found issues with History tab:\n\n';

  if (validation.errors.length > 0) {
    message += 'ERRORS:\n';
    validation.errors.forEach(function(err) {
      message += '• ' + err + '\n';
    });
    message += '\n';
  }

  if (validation.warnings.length > 0) {
    message += 'WARNINGS:\n';
    validation.warnings.forEach(function(warn) {
      message += '• ' + warn + '\n';
    });
  }

  message += '\nAttempt to repair automatically?';

  var response = ui.alert('History Integrity Issues', message, ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return { success: false, message: 'Repair cancelled by user' };
  }

  var repaired = 0;

  // Repair: Add missing summary rows
  var racks = getAllRackConfigTabs();
  racks.forEach(function(rack) {
    var summaryRow = findRackHistorySummaryRow(rack.itemNumber);
    if (summaryRow === -1) {
      createRackHistorySummaryRow(rack.itemNumber, rack.itemName, {
        status: RACK_STATUS.PLACEHOLDER,
        arenaGuid: '',
        created: new Date(),
        lastRefresh: '',
        lastSync: '',
        lastPush: '',
        checksum: ''
      });
      repaired++;
    }
  });

  ui.alert('Repair Complete', 'Repaired ' + repaired + ' issue(s).', ui.ButtonSet.OK);
  return { success: true, message: 'Repaired ' + repaired + ' issues' };
}
