/**
 * BOM Refresh Utilities
 * Handles refreshing rack BOMs from Arena and tracking changes
 */

/**
 * Gets current BOM data from a rack sheet
 * @param {Sheet} sheet - Rack config sheet
 * @return {Array<Object>} Current BOM rows
 */
function getCurrentRackBOMData(sheet) {
  var lastRow = sheet.getLastRow();

  // Row 1 = metadata, Row 2 = headers, Row 3+ = data
  if (lastRow <= 2) {
    return [];
  }

  var dataRange = sheet.getRange(3, 1, lastRow - 2, 6); // Columns A-F (Number, Name, Desc, Category, Lifecycle, Qty)
  var values = dataRange.getValues();

  var bomData = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];

    // Skip empty rows
    if (!row[0]) continue;

    bomData.push({
      rowNumber: i + 3, // Actual sheet row number
      itemNumber: row[0],
      name: row[1],
      description: row[2],
      category: row[3],
      lifecycle: row[4],
      quantity: row[5]
    });
  }

  return bomData;
}

/**
 * Compares current BOM with Arena BOM and detects changes
 * @param {Array<Object>} currentBOM - Current sheet data
 * @param {Array<Object>} arenaBOM - Fresh BOM from Arena
 * @return {Object} Changes categorized as modified, added, removed
 */
function compareBOMs(currentBOM, arenaBOM) {
  var changes = {
    modified: [],
    added: [],
    removed: []
  };

  // Build lookup maps
  var currentMap = {};
  currentBOM.forEach(function(item) {
    currentMap[item.itemNumber] = item;
  });

  var arenaMap = {};
  arenaBOM.forEach(function(line) {
    var bomItem = line.item || line.Item || {};
    var itemNumber = bomItem.number || bomItem.Number || '';
    if (itemNumber) {
      arenaMap[itemNumber] = {
        itemNumber: itemNumber,
        name: bomItem.name || bomItem.Name || '',
        description: bomItem.description || bomItem.Description || '',
        category: (bomItem.category || bomItem.Category || {}).name || (bomItem.category || bomItem.Category || {}).Name || '',
        lifecycle: (bomItem.lifecyclePhase || bomItem.LifecyclePhase || {}).name || (bomItem.lifecyclePhase || bomItem.LifecyclePhase || {}).Name || '',
        quantity: line.quantity || line.Quantity || 1
      };
    }
  });

  // Find modified and added items
  for (var arenaItemNumber in arenaMap) {
    var arenaItem = arenaMap[arenaItemNumber];
    var currentItem = currentMap[arenaItemNumber];

    if (!currentItem) {
      // Item in Arena but not in sheet = ADDED
      changes.added.push(arenaItem);
    } else {
      // Item exists in both - check for modifications
      var itemChanges = [];

      if (currentItem.name !== arenaItem.name) {
        itemChanges.push({
          field: 'Name',
          oldValue: currentItem.name,
          newValue: arenaItem.name
        });
      }

      if (currentItem.description !== arenaItem.description) {
        itemChanges.push({
          field: 'Description',
          oldValue: currentItem.description,
          newValue: arenaItem.description
        });
      }

      if (currentItem.category !== arenaItem.category) {
        itemChanges.push({
          field: 'Category',
          oldValue: currentItem.category,
          newValue: arenaItem.category
        });
      }

      if (currentItem.lifecycle !== arenaItem.lifecycle) {
        itemChanges.push({
          field: 'Lifecycle',
          oldValue: currentItem.lifecycle,
          newValue: arenaItem.lifecycle
        });
      }

      if (currentItem.quantity !== arenaItem.quantity) {
        itemChanges.push({
          field: 'Quantity',
          oldValue: currentItem.quantity,
          newValue: arenaItem.quantity
        });
      }

      if (itemChanges.length > 0) {
        changes.modified.push({
          itemNumber: arenaItemNumber,
          rowNumber: currentItem.rowNumber,
          changes: itemChanges,
          newData: arenaItem
        });
      }
    }
  }

  // Find removed items (in current but not in Arena)
  for (var currentItemNumber in currentMap) {
    if (!arenaMap[currentItemNumber]) {
      changes.removed.push(currentMap[currentItemNumber]);
    }
  }

  return changes;
}

/**
 * Builds a user-friendly summary message of changes
 * @param {Object} changes - Changes object from compareBOMs
 * @return {string} Summary message
 */
function buildChangeSummary(changes) {
  var totalChanges = changes.modified.length + changes.added.length + changes.removed.length;
  var summary = 'Found ' + totalChanges + ' change(s):\n\n';

  if (changes.modified.length > 0) {
    summary += '✏️ ' + changes.modified.length + ' item(s) modified:\n';
    changes.modified.slice(0, 3).forEach(function(mod) {
      summary += '  • ' + mod.itemNumber + ': ' + mod.changes.map(function(c) { return c.field; }).join(', ') + '\n';
    });
    if (changes.modified.length > 3) {
      summary += '  ... and ' + (changes.modified.length - 3) + ' more\n';
    }
    summary += '\n';
  }

  if (changes.added.length > 0) {
    summary += '➕ ' + changes.added.length + ' item(s) added:\n';
    changes.added.slice(0, 3).forEach(function(add) {
      summary += '  • ' + add.itemNumber + ' - ' + add.name + '\n';
    });
    if (changes.added.length > 3) {
      summary += '  ... and ' + (changes.added.length - 3) + ' more\n';
    }
    summary += '\n';
  }

  if (changes.removed.length > 0) {
    summary += '➖ ' + changes.removed.length + ' item(s) removed:\n';
    changes.removed.slice(0, 3).forEach(function(rem) {
      summary += '  • ' + rem.itemNumber + ' - ' + rem.name + '\n';
    });
    if (changes.removed.length > 3) {
      summary += '  ... and ' + (changes.removed.length - 3) + ' more\n';
    }
  }

  return summary;
}

/**
 * Applies BOM changes to the sheet with visual highlighting
 * @param {Sheet} sheet - Rack config sheet
 * @param {Object} changes - Changes to apply
 * @param {string} rackItemNumber - Rack item number for history
 */
function applyBOMChanges(sheet, changes, rackItemNumber) {
  Logger.log('Applying BOM changes...');

  // Track all changes for history
  var historyEntries = [];
  var timestamp = new Date();
  var userName = Session.getActiveUser().getEmail();

  // Apply modifications (with red text)
  changes.modified.forEach(function(mod) {
    var row = mod.rowNumber;

    // Update each changed field
    mod.changes.forEach(function(change) {
      var colIndex;
      var value = change.newValue;

      switch (change.field) {
        case 'Name':
          colIndex = 2; // Column B
          break;
        case 'Description':
          colIndex = 3; // Column C
          break;
        case 'Category':
          colIndex = 4; // Column D
          break;
        case 'Lifecycle':
          colIndex = 5; // Column E
          break;
        case 'Quantity':
          colIndex = 6; // Column F
          break;
      }

      if (colIndex) {
        var cell = sheet.getRange(row, colIndex);
        cell.setValue(value);
        cell.setFontColor('#ff0000'); // Red text
        cell.setBackground('#ffe6e6'); // Light pink background
      }

      // Add to history
      historyEntries.push({
        timestamp: timestamp,
        rack: rackItemNumber,
        itemNumber: mod.itemNumber,
        field: change.field,
        oldValue: change.oldValue,
        newValue: change.newValue,
        changedBy: userName
      });
    });
  });

  // Add new items (at the end)
  if (changes.added.length > 0) {
    var lastRow = sheet.getLastRow();
    var newRowStart = lastRow + 1;

    changes.added.forEach(function(item, index) {
      var row = newRowStart + index;

      sheet.getRange(row, 1).setValue(item.itemNumber);
      sheet.getRange(row, 2).setValue(item.name);
      sheet.getRange(row, 3).setValue(item.description);
      sheet.getRange(row, 4).setValue(item.category);
      sheet.getRange(row, 5).setValue(item.lifecycle);
      sheet.getRange(row, 6).setValue(item.quantity);

      // Blue background for new items
      sheet.getRange(row, 1, 1, 6).setBackground('#d9e7ff');

      // Add to history
      historyEntries.push({
        timestamp: timestamp,
        rack: rackItemNumber,
        itemNumber: item.itemNumber,
        field: 'ADDED',
        oldValue: '',
        newValue: item.name,
        changedBy: userName
      });
    });
  }

  // Mark removed items (strikethrough + gray)
  changes.removed.forEach(function(item) {
    var row = item.rowNumber;
    var range = sheet.getRange(row, 1, 1, 6);
    range.setFontLine('line-through');
    range.setFontColor('#999999');
    range.setBackground('#f5f5f5');

    // Add to history
    historyEntries.push({
      timestamp: timestamp,
      rack: rackItemNumber,
      itemNumber: item.itemNumber,
      field: 'REMOVED',
      oldValue: item.name,
      newValue: '',
      changedBy: userName
    });
  });

  // Write to BOM History tab
  writeBOMHistory(historyEntries);

  Logger.log('BOM changes applied successfully');
}

/**
 * Writes change history to the BOM History tab
 * @param {Array<Object>} entries - History entries to write
 */
function writeBOMHistory(entries) {
  if (entries.length === 0) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = ss.getSheetByName('BOM History');

  // Create sheet if it doesn't exist
  if (!historySheet) {
    historySheet = ss.insertSheet('BOM History');

    // Set up headers
    var headers = ['Timestamp', 'Rack', 'Item Number', 'Field', 'Old Value', 'New Value', 'Changed By'];
    historySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    historySheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1a73e8')
      .setFontColor('white')
      .setFontWeight('bold');

    historySheet.setFrozenRows(1);

    // Set column widths
    historySheet.setColumnWidth(1, 150);  // Timestamp
    historySheet.setColumnWidth(2, 120);  // Rack
    historySheet.setColumnWidth(3, 120);  // Item Number
    historySheet.setColumnWidth(4, 100);  // Field
    historySheet.setColumnWidth(5, 250);  // Old Value
    historySheet.setColumnWidth(6, 250);  // New Value
    historySheet.setColumnWidth(7, 200);  // Changed By

    // Enable text wrapping for Old Value and New Value columns
    historySheet.getRange('E:F').setWrap(true);

    // Set purple tab color to match other system tabs
    historySheet.setTabColor('#9c27b0');

    // Move to end of sheet list
    ss.moveActiveSheet(ss.getNumSheets());
  }

  // Append entries
  var lastRow = historySheet.getLastRow();
  var startRow = lastRow + 1;

  var rows = entries.map(function(entry) {
    return [
      entry.timestamp,
      entry.rack,
      entry.itemNumber,
      entry.field,
      entry.oldValue,
      entry.newValue,
      entry.changedBy
    ];
  });

  historySheet.getRange(startRow, 1, rows.length, 7).setValues(rows);

  Logger.log('Wrote ' + entries.length + ' entries to BOM History');
}

/**
 * Updates the last refreshed timestamp in metadata row
 * @param {Sheet} sheet - Rack config sheet
 */
function updateLastRefreshedTimestamp(sheet) {
  var timestamp = new Date().toLocaleString();

  // Add timestamp to cell E1 (after description)
  sheet.getRange(1, 5).setValue('Last Refreshed: ' + timestamp);
  sheet.getRange(1, 5).setFontStyle('italic').setFontColor('#666666');
}
