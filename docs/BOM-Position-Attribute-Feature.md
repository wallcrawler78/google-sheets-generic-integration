# Rack BOM Location Setting Feature

## Overview

This feature enables automatic tagging of rack items with their position locations when pushing POD structures to Arena. It works similarly to reference designators on electrical components - each rack on a Row's BOM is tagged with its specific positions (e.g., "Pos 1, Pos 3, Pos 8").

**Important:** This feature uses **BOM-level attributes** (not item-level attributes). BOM attributes are configured in Arena and appear on BOM lines, storing information about the relationship between parent and child items.

## Problem Solved

When a Rack appears on a Row's BOM with quantity 5, it means that rack is installed in 5 different positions. Previously, there was no way to track **which specific positions** each rack occupies at the BOM line level. This feature solves that by using a configurable BOM-level attribute to store position information.

## Architecture

### Configuration Storage

Configuration is stored in `PropertiesService.getUserProperties()` with the following keys (defined in `BOMConfiguration.gs`):

- `BOM_POSITION_ATTRIBUTE_GUID` - GUID of the selected Arena BOM attribute
- `BOM_POSITION_ATTRIBUTE_NAME` - Name of the selected attribute (for display)
- `BOM_POSITION_FORMAT` - Format template for position names (default: "Pos {n}")

### Key Components

#### 1. BOMConfiguration.gs

New file containing:
- `showRackBOMLocationSetting()` - Displays configuration dialog
- `loadBOMPositionConfigData()` - Loads available BOM attributes and current settings
- `saveBOMPositionAttribute()` - Saves user's attribute selection
- `getBOMPositionAttributeConfig()` - Retrieves current configuration
- `getBOMAttributes()` - Fetches BOM attributes from Arena via `/settings/items/bom/attributes` (filtered to text types only)

#### 2. ConfigureBOMPositionAttribute.html

UI dialog for selecting and configuring the BOM position attribute:
- Dropdown to select from available Arena BOM attributes (text types only)
- Input field for position name format (with {n} placeholder)
- Live example showing how positions will be formatted
- Option to disable position tracking by selecting "No position tracking"

#### 3. Modified BOMBuilder.gs Functions

**syncBOMToArena() - Lines 324-394**
- Added optional `options` parameter: `{bomAttributes: {itemNumber: attributeValue}}`
- When BOM attributes are provided, includes them in the `additionalAttributes` field of each BOM line payload
- Backward compatible - existing calls without options continue to work

**createRowItems() - Lines 1591-1628**
- Retrieves BOM position attribute configuration via `getBOMPositionAttributeConfig()`
- Builds rack-to-positions mapping from overview scan data
- Formats position names as comma-separated list (e.g., "Pos 1, Pos 3, Pos 8")
- Passes position data to `syncBOMToArena()` via the `bomAttributes` option
- Logs position attribute assignments for debugging

#### 4. Code.gs Menu Integration

Added menu item:
```
Arena Data Center > Configuration > Rack BOM Location Setting
```

## Data Flow

```
1. User configures BOM position attribute via menu
   └─> Selects Arena BOM attribute (text type)
   └─> Configuration stored in PropertiesService

2. User pushes POD structure to Arena
   └─> scanOverviewByRow() extracts rack positions from Overview sheet
   └─> createRowItems() processes each Row:
       ├─> Retrieves position config via getBOMPositionAttributeConfig()
       ├─> Builds rack-to-positions mapping:
       │   └─> {rackItemNumber: ["Pos 1", "Pos 3", "Pos 8"]}
       ├─> Formats as BOM attributes:
       │   └─> {rackItemNumber: {attrGuid: "Pos 1, Pos 3, Pos 8"}}
       └─> Passes to syncBOMToArena() via options.bomAttributes

3. syncBOMToArena() includes position data in BOM line payload
   └─> Arena API: POST /items/{rowGuid}/bom
       └─> Payload includes additionalAttributes with position values
```

## Usage Instructions

### Setup (One-Time)

1. **Create BOM attribute in Arena** (if not exists):
   - Go to Arena PLM Settings
   - Create custom BOM attribute (type: SINGLE_LINE_TEXT)
   - Name suggestion: "Rack Positions" or "Position Designators"
   - Apply to relevant categories (e.g., Row items)

2. **Configure in Sheets**:
   - Open your Data Center spreadsheet
   - Menu: `Arena Data Center > Configuration > Configure BOM Position Attribute`
   - Select the BOM attribute from the dropdown
   - Optionally adjust position format (default: "Pos {n}")
   - Click "Save Configuration"

### Using the Feature

Once configured, the feature works automatically when you push POD structures:

1. Menu: `Arena Data Center > BOM Operations > Push POD Structure to Arena`
2. The system will:
   - Scan Overview sheet for rack positions
   - Create Row items in Arena
   - Automatically tag each rack with its positions in the configured BOM attribute
   - Create POD item with Row BOM

### Disabling Position Tracking

To disable position tracking:
- Menu: `Arena Data Center > Configuration > Configure BOM Position Attribute`
- Select "-- No position tracking (disabled) --"
- Click "Save Configuration"

## Example

**Overview Sheet:**
```
Row | Pos 1      | Pos 2       | Pos 3      | ... | Pos 8
1   | RACK-001   | (empty)     | RACK-001   | ... | RACK-002
2   | RACK-003   | RACK-003    | RACK-003   | ... | (empty)
```

**Resulting Arena BOM (for Row 1):**

| Item Number | Qty | Rack Positions (BOM Attribute) |
|-------------|-----|-------------------------------|
| RACK-001    | 2   | Pos 1, Pos 3                  |
| RACK-002    | 1   | Pos 8                         |

**Resulting Arena BOM (for Row 2):**

| Item Number | Qty | Rack Positions (BOM Attribute) |
|-------------|-----|-------------------------------|
| RACK-003    | 3   | Pos 1, Pos 2, Pos 3           |

## Technical Details

### Arena API Payload Structure

When creating a BOM line with position attributes:

```javascript
{
  "item": {
    "guid": "ABC123..."
  },
  "quantity": 2,
  "level": 0,
  "lineNumber": 1,
  "additionalAttributes": {
    "ATTR-GUID-HERE": "Pos 1, Pos 3"
  }
}
```

### Position Name Detection

Position names are auto-detected from Overview sheet column headers:
- Looks for columns starting with "pos" (case-insensitive)
- Extracts exact header text (e.g., "Pos 1", "Position 3", etc.)
- No hardcoded position naming - adapts to your sheet structure

### Attribute Type Filtering

The configuration UI only shows text-type attributes:
- `SINGLE_LINE_TEXT`
- `MULTI_LINE_TEXT`
- `FIXED_DROP_DOWN`

Checkboxes, numbers, and other non-text types are filtered out since they can't store position lists.

## Benefits

✅ **Flexible** - User picks which BOM attribute to use
✅ **Scalable** - Position names auto-detected from sheet headers
✅ **Automatic** - System builds mapping from existing overview data
✅ **Non-breaking** - Only applies when configured; existing flows unchanged
✅ **Intuitive** - Mirrors reference designator pattern users already understand

## Troubleshooting

### Error: "The additional attribute is not recognized"

This error occurs when the configured attribute is an **Item attribute** instead of a **BOM attribute**.

**Solution:**
1. Open Arena Data Center > Configuration > Rack BOM Location Setting
2. Clear the current configuration (if shown as invalid)
3. Select a valid BOM-level attribute from the dropdown
4. Verify in Arena that the attribute appears in your Custom BOM views (not just item specs)

**How to verify BOM vs Item attributes:**
- **BOM attributes**: Appear in Arena under Items > Custom BOMs > Attribute column
- **Item attributes**: Appear on item specs page (Name, Description, etc.)
- **API difference**: Fetched from `/settings/items/bom/attributes` (BOM) vs `/settings/items/attributes` (Item)

### Configuration shows warning about "Previously configured attribute not found"

This means the attribute was deleted from Arena, or was an Item attribute mistakenly configured before the validation was added.

**Solution:**
1. The system automatically clears invalid configurations
2. Select a new valid BOM attribute from the dropdown
3. Save the configuration

### How to create a BOM attribute in Arena

If you don't have a suitable BOM attribute:

1. In Arena, go to **Workspaces > Items > Custom BOMs**
2. Create or edit a Custom BOM view for your category (e.g., "POD", "ROW")
3. Add an attribute column (Arena will show you BOM-level attributes)
4. Create a new BOM attribute if needed:
   - Type: Single Line Text or Multi Line Text
   - Name: e.g., "Rack Location" or "Position"
5. Return to Google Sheets and refresh the Rack BOM Location Setting dialog
6. Select your new BOM attribute

### POD push fails with BOM line errors

**Symptoms:**
- Error during Row creation: "Failed to add BOM line..."
- Attribute-related error messages

**Solutions:**
1. Check that attribute is configured for the ROW category's Custom BOM view
2. Verify attribute GUID matches between configuration and Arena
3. Try clearing configuration and running POD push without position tracking
4. Reconfigure with a known-good BOM attribute

## Future Enhancements

Potential improvements:
1. **Custom formatting**: Support more complex position formats (e.g., "Row 1, Pos 3")
2. **Multiple attributes**: Support additional BOM-level attributes (notes, find numbers, etc.)
3. **Validation**: Pre-check that selected attribute is configured for the BOM category
4. **Bulk operations**: Apply position tracking to existing POD structures retroactively

## Files Modified

- **NEW:** `BOMConfiguration.gs` - Configuration management functions
- **NEW:** `ConfigureBOMPositionAttribute.html` - Configuration UI
- **NEW:** `Docs/BOM-Position-Attribute-Feature.md` - This documentation
- **MODIFIED:** `BOMBuilder.gs`
  - `syncBOMToArena()` - Added options parameter for BOM attributes
  - `createRowItems()` - Added position mapping logic
- **MODIFIED:** `Code.gs` - Added menu item for configuration

## Testing Recommendations

1. **Test with no configuration**: Verify POD push works normally when position attribute is not configured
2. **Test with configuration**: Configure position attribute and verify BOM lines include position data
3. **Test multiple racks in same positions**: Verify positions are correctly aggregated
4. **Test Arena API errors**: Verify graceful handling if attribute doesn't exist on category
5. **Test position format**: Verify custom position formats work correctly

## Related Documentation

- Arena API BOM endpoints: https://api.arenasolutions.com/v1/documentation
- Google Apps Script PropertiesService: https://developers.google.com/apps-script/reference/properties
