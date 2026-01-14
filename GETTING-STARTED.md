# Getting Started - Installation Guide

Welcome! This guide will help you install the Arena Sheets Integration into your Google Sheets spreadsheet using the command line.

## What You'll Need

- A Google account
- A terminal/command prompt
- Node.js installed on your computer ([Download here](https://nodejs.org/))
- 10 minutes of your time

---

## Step 1: Install clasp (Google's Command Line Tool)

Open your terminal and run:

```bash
npm install -g @google/clasp
```

**What is clasp?** It's Google's official tool for managing Google Apps Script projects from your computer.

---

## Step 2: Log in to Google

Authenticate clasp with your Google account:

```bash
clasp login
```

This will open your browser. Sign in with the Google account you want to use for the spreadsheet.

---

## Step 3: Create a New Google Sheets Spreadsheet

Create a new spreadsheet with an attached Apps Script project:

```bash
clasp create --type sheets --title "Arena Integration"
```

**What this does:**
- Creates a new Google Sheets spreadsheet named "Arena Integration"
- Creates an Apps Script project attached to that spreadsheet
- Creates a `.clasp.json` file in your current directory linking them together

**Note:** If you already have a spreadsheet and want to use that instead, see the [Advanced Options](#advanced-options) section below.

---

## Step 4: Push the Code to Your Spreadsheet

Navigate to the project directory (if not already there):

```bash
cd /path/to/PTC-Arena-Sheets-DataCenter
```

Push all the code files to Google Apps Script:

```bash
clasp push
```

**What this does:**
- Uploads all `.gs` (Google Apps Script) files
- Uploads all `.html` files
- Uploads the `appsscript.json` configuration file

You should see: `Pushed 41 files.`

---

## Step 5: Open Your Spreadsheet

Open the spreadsheet in your browser:

```bash
clasp open
```

This will open the Google Sheets spreadsheet. Alternatively, find it in your Google Drive under "Arena Integration".

---

## Step 6: First Time Setup

When you open the spreadsheet for the first time:

### For New Users (Fresh Spreadsheet):
1. The **Setup Wizard** will appear automatically
2. Follow the wizard to configure your terminology and types
3. Choose what you want to call your primary entity (e.g., "Rack", "Server", "Product")
4. Define your type classifications with keywords
5. Complete the wizard to create your initial sheet structure

### For Existing Datacenter Users:
1. The system will **auto-detect** your datacenter configuration
2. **Auto-migration** runs silently in the background
3. A one-time notification explains the new configuration system
4. Everything continues to work exactly as before
5. You can customize terminology later via **Setup → Configure Type System**

---

## Step 7: Authorize the Application

The first time you use a feature, Google will ask for permissions:

1. Click **Continue** when prompted
2. Select your Google account
3. Click **Advanced** → **Go to Arena Integration (unsafe)**
4. Review permissions and click **Allow**

**Why does it say "unsafe"?** Because this is an unverified custom script. You're installing your own code, so this is expected and safe.

---

## You're Ready to Go!

Your Arena Sheets Integration is now installed and configured.

### Next Steps:
- Read the [Documentation](./Docs/README.md) for full feature guide
- Configure your Arena API connection via **Arena → Settings → Configure API**
- Start adding items via **Arena → Item Picker**
- Create your first configuration sheet

---

## Advanced Options

### Option A: Use an Existing Spreadsheet

If you already have a spreadsheet and want to add this code to it:

1. Open your spreadsheet in Google Sheets
2. Go to **Extensions → Apps Script**
3. Note the Script ID from the URL: `https://script.google.com/.../**SCRIPT_ID**/edit`
4. In your terminal, run:
   ```bash
   clasp clone SCRIPT_ID
   ```
5. Copy all `.gs` and `.html` files from this project into the cloned directory
6. Run `clasp push`

### Option B: Deploy to Multiple Spreadsheets

To use this code in multiple spreadsheets:

1. Complete the initial setup (Steps 1-4)
2. For each additional spreadsheet:
   ```bash
   clasp create --type sheets --title "Arena Integration 2"
   clasp push
   ```
3. Each spreadsheet will have its own configuration

### Option C: Development Workflow

If you're modifying the code:

1. Make changes to `.gs` or `.html` files locally
2. Push changes: `clasp push`
3. Open and test: `clasp open`
4. Pull changes from Google (if edited online): `clasp pull`

**Watch mode** (auto-push on file changes):
```bash
clasp push --watch
```

---

## Troubleshooting

### "clasp: command not found"
**Solution:** Install clasp globally with npm:
```bash
npm install -g @google/clasp
```

### "User has not enabled the Apps Script API"
**Solution:** Enable the API at https://script.google.com/home/usersettings

### "Push failed: Invalid argument"
**Solution:** Make sure you're in the correct directory with `.clasp.json` file

### "Authorization required"
**Solution:** Run `clasp login` again to re-authenticate

### Setup Wizard shows an error
**Solution:** The HTML wizard UI is not created yet (Phase 4 work). Datacenter users can ignore this - auto-migration handles setup. New users will need to wait for the full wizard implementation.

---

## Project Structure

After installation, your Google Apps Script project contains:

```
📦 Arena Integration (Google Sheets)
 ┣ 📜 Code.gs              - Main entry point, menu system
 ┣ 📜 TypeSystemConfig.gs  - Configuration management
 ┣ 📜 MigrationManager.gs  - Auto-migration system
 ┣ 📜 Config.gs            - Type classification logic
 ┣ 📜 ArenaAPI.gs          - Arena PLM API integration
 ┣ 📜 BOMBuilder.gs        - Bill of Materials operations
 ┣ 📜 ItemPicker.html      - Item selection dialog
 ┣ 📜 ... and 34 more files
```

---

## Getting Help

- **Documentation**: [Docs/README.md](./Docs/README.md)
- **Architecture**: [Docs/ARCHITECTURE.md](./Docs/ARCHITECTURE.md)
- **Security Audit**: [Docs/SECURITY-AUDIT.md](./Docs/SECURITY-AUDIT.md)
- **Configuration Guide**: Coming soon in Phase 5
- **Issues**: [GitHub Issues](https://github.com/wallcrawler78/PTC-Arena-Sheets-DataCenter/issues)

---

## What's Next?

Once installed, explore these features:

1. **Item Picker** - Search and insert Arena items
2. **Type Classification** - Automatically categorize items by keywords
3. **BOM Builder** - Generate Bills of Materials
4. **Hierarchy Push** - Push structures back to Arena
5. **Configuration System** - Customize all terminology
6. **History Tracking** - Track all changes with timestamps

Happy integrating! 🚀
