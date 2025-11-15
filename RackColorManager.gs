/**
 * Rack Color Manager
 * Handles color assignments for rack configurations in overview sheets
 */

// Script Properties key for rack color configuration
var PROP_RACK_COLORS = 'RACK_COLORS';

/**
 * Gets rack color configuration (custom overrides only)
 * @return {Object} Map of rack number to custom color
 */
function getRackColors() {
  var json = PropertiesService.getScriptProperties().getProperty(PROP_RACK_COLORS);
  if (!json) {
    return {};
  }
  try {
    return JSON.parse(json);
  } catch (e) {
    Logger.log('Error parsing rack colors: ' + e.message);
    return {};
  }
}

/**
 * Saves rack color configuration
 * @param {Object} colors - Map of rack number to color
 */
function saveRackColors(colors) {
  PropertiesService.getScriptProperties().setProperty(
    PROP_RACK_COLORS,
    JSON.stringify(colors)
  );
}

/**
 * Gets color for a specific rack (custom or auto-generated)
 * @param {string} rackNumber - Rack item number
 * @return {string} Hex color code
 */
function getRackColor(rackNumber) {
  var customColors = getRackColors();

  // Return custom color if set
  if (customColors[rackNumber]) {
    return customColors[rackNumber];
  }

  // Otherwise return auto-generated color
  return getDefaultRackColor(rackNumber);
}

/**
 * Sets a custom color for a specific rack
 * @param {string} rackNumber - Rack item number
 * @param {string} color - Hex color code
 */
function setRackColor(rackNumber, color) {
  var colors = getRackColors();
  colors[rackNumber] = color;
  saveRackColors(colors);
}

/**
 * Removes custom color for a rack (reverts to auto color)
 * @param {string} rackNumber - Rack item number
 */
function clearRackColor(rackNumber) {
  var colors = getRackColors();
  delete colors[rackNumber];
  saveRackColors(colors);
}

/**
 * Generates a consistent auto color for a rack based on its item number
 * Uses hash-based HSL color generation for pastel shades
 * @param {string} rackNumber - Rack item number
 * @return {string} Hex color code
 */
function getDefaultRackColor(rackNumber) {
  if (!rackNumber) {
    return '#E3F2FD'; // Default light blue
  }

  // Simple hash function for consistent colors
  var hash = 0;
  for (var i = 0; i < rackNumber.length; i++) {
    hash = rackNumber.charCodeAt(i) + ((hash << 5) - hash);
  }

  // Convert to HSL for better color distribution
  var h = Math.abs(hash % 360);           // Hue: 0-360
  var s = 65 + (Math.abs(hash) % 20);     // Saturation: 65-85%
  var l = 80 + (Math.abs(hash >> 8) % 15); // Lightness: 80-95% (pastel)

  return hslToHex(h, s, l);
}

/**
 * Converts HSL color to Hex
 * @param {number} h - Hue (0-360)
 * @param {number} s - Saturation (0-100)
 * @param {number} l - Lightness (0-100)
 * @return {string} Hex color code
 */
function hslToHex(h, s, l) {
  s /= 100;
  l /= 100;

  var c = (1 - Math.abs(2 * l - 1)) * s;
  var x = c * (1 - Math.abs((h / 60) % 2 - 1));
  var m = l - c/2;
  var r = 0, g = 0, b = 0;

  if (0 <= h && h < 60) {
    r = c; g = x; b = 0;
  } else if (60 <= h && h < 120) {
    r = x; g = c; b = 0;
  } else if (120 <= h && h < 180) {
    r = 0; g = c; b = x;
  } else if (180 <= h && h < 240) {
    r = 0; g = x; b = c;
  } else if (240 <= h && h < 300) {
    r = x; g = 0; b = c;
  } else if (300 <= h && h < 360) {
    r = c; g = 0; b = x;
  }

  var rr = Math.round((r + m) * 255);
  var gg = Math.round((g + m) * 255);
  var bb = Math.round((b + m) * 255);

  // Convert to hex with zero padding
  var rHex = rr.toString(16);
  if (rHex.length === 1) rHex = '0' + rHex;

  var gHex = gg.toString(16);
  if (gHex.length === 1) gHex = '0' + gHex;

  var bHex = bb.toString(16);
  if (bHex.length === 1) bHex = '0' + bHex;

  return '#' + rHex + gHex + bHex;
}

/**
 * Loads rack color data for the configuration UI
 * @return {Object} Object containing racks and colors
 */
function loadRackColorData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = spreadsheet.getSheets();
  var racks = [];

  // Find all rack configuration sheets
  allSheets.forEach(function(sheet) {
    var sheetName = sheet.getName().toLowerCase();

    // Skip non-rack sheets
    if (sheetName.indexOf('overview') !== -1 ||
        sheetName.indexOf('legend') !== -1 ||
        sheetName.indexOf('config') !== -1 ||
        sheetName.indexOf('bom') !== -1) {
      return;
    }

    // Look for rack pattern: item number in name
    var match = sheet.getName().match(/(\d{3}-\d{4})/);
    if (match) {
      var itemNumber = match[1];

      // Try to get item name from sheet
      var itemName = '';
      try {
        var nameCell = sheet.getRange('B1');
        var nameValue = nameCell.getValue();
        if (nameValue && typeof nameValue === 'string') {
          itemName = nameValue.replace(/^Name:\s*/i, '').trim();
        }
      } catch (e) {
        // Ignore errors reading name
      }

      racks.push({
        itemNumber: itemNumber,
        itemName: itemName,
        sheetName: sheet.getName()
      });
    }
  });

  // Sort racks by item number
  racks.sort(function(a, b) {
    return a.itemNumber.localeCompare(b.itemNumber);
  });

  return {
    racks: racks,
    colors: getRackColors()
  };
}
