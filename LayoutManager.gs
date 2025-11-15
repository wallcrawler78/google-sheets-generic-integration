/**
 * Layout Manager
 * Handles creation and management of different sheet layouts (Tower, Overview, Rack Config)
 */

/**
 * Creates a tower layout sheet (vertical server stacking)
 * @param {string} sheetName - Name for the new tower sheet
 * @return {Sheet} The created sheet
 */
function createTowerLayout(sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Check if sheet already exists
  var existingSheet = spreadsheet.getSheetByName(sheetName);
  if (existingSheet) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      'Sheet Exists',
      'A sheet named "' + sheetName + '" already exists. Overwrite?',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      spreadsheet.deleteSheet(existingSheet);
    } else {
      return null;
    }
  }

  // Create new sheet
  var sheet = spreadsheet.insertSheet(sheetName);

  // Set up headers
  var headers = ['Position', 'Qty', 'Item Number', 'Item Name', 'Category', 'Notes'];

  // Add configured attribute columns
  var columns = getItemColumns();
  columns.forEach(function(col) {
    headers.push(col.header || col.attributeName);
  });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 80);  // Position
  sheet.setColumnWidth(2, 60);  // Qty
  sheet.setColumnWidth(3, 150); // Item Number
  sheet.setColumnWidth(4, 250); // Item Name
  sheet.setColumnWidth(5, 120); // Category
  sheet.setColumnWidth(6, 200); // Notes

  // Add initial rows (positions)
  var positions = [];
  for (var i = 1; i <= 42; i++) {
    positions.push(['U' + i, 1, '', '', '', '']);
  }

  if (positions.length > 0) {
    sheet.getRange(2, 1, positions.length, 6).setValues(positions);
  }

  // Freeze header row and position column
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  // Add alternating row colors for better readability
  for (var row = 2; row <= positions.length + 1; row++) {
    if (row % 2 === 0) {
      sheet.getRange(row, 1, 1, headers.length).setBackground('#f8f9fa');
    }
  }

  return sheet;
}

/**
 * Creates an overview layout sheet (horizontal rack grid)
 * @param {string} sheetName - Name for the new overview sheet
 * @param {number} rows - Number of rows in the grid
 * @param {number} cols - Number of columns in the grid
 * @return {Sheet} The created sheet
 */
function createOverviewLayout(sheetName, rows, cols) {
  rows = rows || 10;
  cols = cols || 10;

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Check if sheet already exists
  var existingSheet = spreadsheet.getSheetByName(sheetName);
  if (existingSheet) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      'Sheet Exists',
      'A sheet named "' + sheetName + '" already exists. Overwrite?',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      spreadsheet.deleteSheet(existingSheet);
    } else {
      return null;
    }
  }

  // Create new sheet
  var sheet = spreadsheet.insertSheet(sheetName);

  // Set up title
  sheet.getRange(1, 1).setValue('Datacenter Overview');
  sheet.getRange(1, 1, 1, cols).merge()
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff');

  // Create grid
  var startRow = 3;

  // Add column headers (Pos 1, Pos 2, Pos 3, etc.)
  var colHeaders = [];
  for (var c = 0; c < cols; c++) {
    colHeaders.push('Pos ' + (c + 1));
  }
  sheet.getRange(startRow, 2, 1, cols).setValues([colHeaders]);
  sheet.getRange(startRow, 2, 1, cols)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#f0f0f0');

  // Add row headers (1, 2, 3, etc.)
  var rowHeaders = [];
  for (var r = 0; r < rows; r++) {
    rowHeaders.push([r + 1]);
  }
  sheet.getRange(startRow + 1, 1, rows, 1).setValues(rowHeaders);
  sheet.getRange(startRow + 1, 1, rows, 1)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#f0f0f0');

  // Set cell sizes for grid
  sheet.setColumnWidth(1, 50); // Row header column
  for (var c = 2; c <= cols + 1; c++) {
    sheet.setColumnWidth(c, 120);
  }

  for (var r = startRow + 1; r <= startRow + rows; r++) {
    sheet.setRowHeight(r, 80);
  }

  // Add borders to grid
  var gridRange = sheet.getRange(startRow + 1, 2, rows, cols);
  gridRange.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  // Freeze headers (only freeze rows, not columns due to merged title)
  sheet.setFrozenRows(startRow);

  return sheet;
}

/**
 * Creates a rack configuration sheet with standard BOM structure
 * @param {string} rackName - Name of the rack (e.g., "Rack A")
 * @return {Sheet} The created sheet
 */
function createRackConfigSheet(rackName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Check if sheet already exists
  var existingSheet = spreadsheet.getSheetByName(rackName);
  if (existingSheet) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      'Sheet Exists',
      'A sheet named "' + rackName + '" already exists. Overwrite?',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      spreadsheet.deleteSheet(existingSheet);
    } else {
      return null;
    }
  }

  // Create new sheet
  var sheet = spreadsheet.insertSheet(rackName);

  // Set up headers for BOM structure
  var headers = ['Level', 'Qty', 'Item Number', 'Item Name', 'Category', 'Lifecycle', 'Notes'];

  // Add configured attribute columns
  var columns = getItemColumns();
  columns.forEach(function(col) {
    headers.push(col.header || col.attributeName);
  });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 60);  // Level
  sheet.setColumnWidth(2, 60);  // Qty
  sheet.setColumnWidth(3, 150); // Item Number
  sheet.setColumnWidth(4, 250); // Item Name
  sheet.setColumnWidth(5, 120); // Category
  sheet.setColumnWidth(6, 100); // Lifecycle
  sheet.setColumnWidth(7, 200); // Notes

  // Add rack info section at top
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1).setValue('Rack Configuration: ' + rackName)
    .setFontSize(14)
    .setFontWeight('bold');

  sheet.getRange(1, 1, 1, headers.length).merge()
    .setBackground('#f8f9fa');

  // Freeze header row
  sheet.setFrozenRows(2);

  return sheet;
}

/**
 * Links a cell in the overview to a rack configuration sheet
 * @param {string} overviewSheetName - Name of the overview sheet
 * @param {number} row - Row in overview grid
 * @param {number} col - Column in overview grid
 * @param {string} rackSheetName - Name of the rack sheet to link to
 */
function linkOverviewToRack(overviewSheetName, row, col, rackSheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var overviewSheet = spreadsheet.getSheetByName(overviewSheetName);

  if (!overviewSheet) {
    throw new Error('Overview sheet not found: ' + overviewSheetName);
  }

  var cell = overviewSheet.getRange(row, col);

  // Create hyperlink formula
  var sheetId = spreadsheet.getSheetByName(rackSheetName).getSheetId();
  var url = '#gid=' + sheetId;

  // Set formula with rack name
  cell.setFormula('=HYPERLINK("' + url + '", "' + rackSheetName + '")');
  cell.setFontColor('#1a73e8')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

/**
 * Populates an overview grid cell with rack information
 * @param {string} overviewSheetName - Name of the overview sheet
 * @param {number} row - Row in overview grid
 * @param {number} col - Column in overview grid
 * @param {string} rackName - Rack name/identifier
 * @param {string} category - Rack category for color coding
 */
function populateOverviewCell(overviewSheetName, row, col, rackName, category) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var overviewSheet = spreadsheet.getSheetByName(overviewSheetName);

  if (!overviewSheet) {
    throw new Error('Overview sheet not found: ' + overviewSheetName);
  }

  var cell = overviewSheet.getRange(row, col);
  cell.setValue(rackName);

  // Apply category color
  var color = getCategoryColor(category);
  if (color) {
    cell.setBackground(color);
  }

  // Format
  cell.setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontWeight('bold');
}

/**
 * Menu action to create a new tower layout
 */
function createNewTowerLayout() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt(
    'Create Tower Layout',
    'Enter name for the tower layout (e.g., "Tower A"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    var name = response.getResponseText().trim();
    if (!name) {
      ui.alert('Error', 'Tower name is required', ui.ButtonSet.OK);
      return;
    }

    try {
      var sheet = createTowerLayout(name);
      if (sheet) {
        ui.alert('Success', 'Tower layout "' + name + '" created!', ui.ButtonSet.OK);
        SpreadsheetApp.setActiveSheet(sheet);
      }
    } catch (error) {
      ui.alert('Error', 'Failed to create tower layout: ' + error.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * Menu action to create a new overview layout
 */
function createNewOverviewLayout() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt(
    'Create Overview Layout',
    'Enter name for the overview (e.g., "Hall 1 Overview"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    var name = response.getResponseText().trim();
    if (!name) {
      ui.alert('Error', 'Overview name is required', ui.ButtonSet.OK);
      return;
    }

    // Prompt for number of rows
    var rowsResponse = ui.prompt(
      'Number of Rows',
      'Enter number of rows in the datacenter (e.g., "10"):',
      ui.ButtonSet.OK_CANCEL
    );

    if (rowsResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    var rows = parseInt(rowsResponse.getResponseText().trim(), 10);
    if (isNaN(rows) || rows < 1 || rows > 50) {
      ui.alert('Error', 'Number of rows must be between 1 and 50', ui.ButtonSet.OK);
      return;
    }

    // Prompt for number of rack positions per row
    var positionsResponse = ui.prompt(
      'Rack Positions per Row',
      'Enter number of rack positions per row (e.g., "12"):',
      ui.ButtonSet.OK_CANCEL
    );

    if (positionsResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    var positions = parseInt(positionsResponse.getResponseText().trim(), 10);
    if (isNaN(positions) || positions < 1 || positions > 50) {
      ui.alert('Error', 'Number of positions must be between 1 and 50', ui.ButtonSet.OK);
      return;
    }

    try {
      var sheet = createOverviewLayout(name, rows, positions);
      if (sheet) {
        ui.alert('Success',
          'Overview layout "' + name + '" created!\n\n' +
          'Rows: ' + rows + '\n' +
          'Positions per row: ' + positions + '\n\n' +
          'Use "Show Rack Picker" to place racks in the grid.',
          ui.ButtonSet.OK);
        SpreadsheetApp.setActiveSheet(sheet);
      }
    } catch (error) {
      ui.alert('Error', 'Failed to create overview layout: ' + error.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * Menu action to create a new rack configuration sheet
 */
function createNewRackConfig() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt(
    'Create Rack Configuration',
    'Enter rack name (e.g., "Rack A"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    var name = response.getResponseText().trim();
    if (!name) {
      ui.alert('Error', 'Rack name is required', ui.ButtonSet.OK);
      return;
    }

    try {
      var sheet = createRackConfigSheet(name);
      if (sheet) {
        ui.alert('Success', 'Rack configuration "' + name + '" created!', ui.ButtonSet.OK);
        SpreadsheetApp.setActiveSheet(sheet);
      }
    } catch (error) {
      ui.alert('Error', 'Failed to create rack configuration: ' + error.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * Auto-links all rack sheets to the overview layout
 * @param {string} overviewSheetName - Name of the overview sheet
 */
function autoLinkRacksToOverview(overviewSheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = spreadsheet.getSheets();

  // Find all rack sheets
  var rackSheets = [];
  allSheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (name.toLowerCase().indexOf('rack') !== -1 && name !== overviewSheetName) {
      rackSheets.push(name);
    }
  });

  if (rackSheets.length === 0) {
    SpreadsheetApp.getUi().alert('No Racks', 'No rack sheets found to link', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Link racks to overview grid (arrange in grid pattern)
  var startRow = 4; // After headers
  var startCol = 2;
  var maxCols = 5; // 5 racks per row

  rackSheets.forEach(function(rackName, index) {
    var row = startRow + Math.floor(index / maxCols);
    var col = startCol + (index % maxCols);

    try {
      linkOverviewToRack(overviewSheetName, row, col, rackName);
    } catch (error) {
      Logger.log('Error linking rack ' + rackName + ': ' + error.message);
    }
  });

  SpreadsheetApp.getUi().alert(
    'Success',
    'Linked ' + rackSheets.length + ' rack sheets to overview',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Menu action to auto-link racks
 */
function autoLinkRacksToOverviewAction() {
  var ui = SpreadsheetApp.getUi();

  // Find overview sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = spreadsheet.getSheets();
  var overviewSheets = [];

  allSheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (name.toLowerCase().indexOf('overview') !== -1) {
      overviewSheets.push(name);
    }
  });

  if (overviewSheets.length === 0) {
    ui.alert('No Overview', 'No overview sheet found. Create one first.', ui.ButtonSet.OK);
    return;
  }

  var overviewName = overviewSheets[0];
  if (overviewSheets.length > 1) {
    var response = ui.prompt(
      'Multiple Overviews Found',
      'Enter the name of the overview sheet to use:\n' + overviewSheets.join(', '),
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() === ui.Button.OK) {
      overviewName = response.getResponseText().trim();
    } else {
      return;
    }
  }

  autoLinkRacksToOverview(overviewName);
}
