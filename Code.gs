// Global constants for internal sheet names
const BORROW_SHEET_NAME = "Borrow Tools";
const MASTER_SHEET_NAME = "ToExcel_MTL_AssetManagementTable";

// ============================================================================
// CONFIGURATION: REAL-TIME JOB DATABASE (DIRECT DRIVE & FALLBACK CSVs)
// ============================================================================
const EXTERNAL_JOB_DB_ID = '1vGPJvUOgGu7xEehsXu82QFM04qdo513pW8r3XFnzJRM';
const EXTERNAL_JOB_DB_SHEET_NAMES = ['OOR', 'New Orders', 'STOCK ITEMS'];

// Published CSV endpoints serving as fallbacks if a user lacks direct Drive permissions
const JOB_DB_CSV_FALLBACKS = [
  {
    name: "OOR",
    gid: "1402212990",
    url: "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVbefWc5DMF-QD7sRDkZfu6-pyEsbkqiq2uekx_0d_zREebFp7tk3BPeqam3HBh3xsT60sXbc2G1hj/pub?gid=1402212990&single=true&output=csv"
  },
  {
    name: "New Orders",
    gid: "1609933687",
    url: "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVbefWc5DMF-QD7sRDkZfu6-pyEsbkqiq2uekx_0d_zREebFp7tk3BPeqam3HBh3xsT60sXbc2G1hj/pub?gid=1609933687&single=true&output=csv"
  },
  {
    name: "STOCK ITEMS",
    gid: "1667615887",
    url: "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVbefWc5DMF-QD7sRDkZfu6-pyEsbkqiq2uekx_0d_zREebFp7tk3BPeqam3HBh3xsT60sXbc2G1hj/pub?gid=1667615887&single=true&output=csv"
  }
];

/**
 * Creates custom spreadsheet menu on document open
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Asset Management')
      .addItem('Open Main Menu', 'showMainMenuDialog')
      .addSeparator()
      .addItem('Import New Assets', 'showImportDialog')
      .addSeparator()
      .addItem('Sync External Job Tabs (OOR / New Orders / Stock)', 'syncJobDatabaseTabs')
      .addToUi();
}

/**
 * Modal display handlers for HTML dialogs
 */
function showMainMenuDialog() {
  const html = HtmlService.createHtmlOutputFromFile('MainMenu').setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Asset Management Menu');
}

function showBorrowDialog() {
  const html = HtmlService.createHtmlOutputFromFile('BorrowDialog').setWidth(700).setHeight(650); 
  SpreadsheetApp.getUi().showModalDialog(html, 'Borrow Asset');
}

function showReturnDialog(assetId) {
  const template = HtmlService.createTemplateFromFile('ReturnDialog');
  template.assetId = assetId || '';
  const html = template.evaluate().setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Return Asset');
}

function showFindDialog() {
  const html = HtmlService.createHtmlOutputFromFile('FindDialog').setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Find Asset');
}

function showImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ImportDialog').setWidth(700).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import New Assets');
}

/**
 * Dynamic Column Index Mapping Helper for Master Asset Sheet
 * Normalizes header strings and returns zero-based column indices.
 */
function getColumnIndices(headerRow) {
  if (!headerRow || !Array.isArray(headerRow)) {
    return {
      partNum: -1, desc: -1, assetId: -1, extDesc: -1, location: -1,
      assignedTo: -1, status: -1, category: -1, group: -1, mfg: -1
    };
  }

  const normalized = headerRow.map(h => h ? h.toString().trim().toUpperCase() : "");

  const find = (possibleNames) => {
    for (const name of possibleNames) {
      const idx = normalized.indexOf(name);
      if (idx !== -1) return idx;
    }
    return -1;
  };

  return {
    partNum: find(["PART NUMBER", "PART NO", "ITEM", "PART"]),
    desc: find(["DESCRIPTION", "DESC"]),
    assetId: find(["ASSET", "ASSET ID", "TAG", "ASSET#"]),
    extDesc: find(["EXT. DESCRIPTION", "EXTENDED DESCRIPTION", "DETAILS"]),
    location: find(["LOCATION", "LOC", "BIN", "SHELF", "STOCKROOM LOCATION"]),
    assignedTo: find(["ASSIGNED TO", "BORROWER", "CHECKED OUT TO", "PC"]),
    status: find(["STATUS", "STATE"]),
    category: find(["CATEGORY", "CLASS"]),
    group: find(["GROUP", "TYPE"]),
    mfg: find(["MANUFACTURER", "MFG", "BRAND", "PRODUCT CLASS"])
  };
}

/**
 * 100% Real-Time Multi-Source Job Details Resolver
 * 
 * Order of Operation:
 *  1. Direct Google Drive Sheet Access (0-second delay; instant updates)
 *     Searches 'OOR', 'New Orders', and 'STOCK ITEMS' in real time.
 *  2. Local Workbook Sheets (if user copied/synced tabs locally)
 *  3. Fallback to Published CSV URLs (if the scanning user lacks Drive access)
 */
function getJobDetails(jobOrder) {
  if (!jobOrder) return null;
  const cleanJob = jobOrder.toString().trim().toUpperCase();

  // Tier 1: 100% REAL-TIME DIRECT ACCESS via SpreadsheetApp
  try {
    const extSs = SpreadsheetApp.openById(EXTERNAL_JOB_DB_ID);
    for (const sheetName of EXTERNAL_JOB_DB_SHEET_NAMES) {
      const sheet = extSs.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) continue;

      const result = searchSheetForJob(sheet, cleanJob, sheetName);
      if (result && result.found) {
        return result;
      }
    }
  } catch (driveErr) {
    console.warn(`Direct Google Drive lookup skipped/inaccessible: ${driveErr}`);
  }

  // Tier 2: Check Local Tabs if synced into the active workbook
  const activeSs = SpreadsheetApp.getActiveSpreadsheet();
  for (const sheetName of EXTERNAL_JOB_DB_SHEET_NAMES) {
    const localSheet = activeSs.getSheetByName(sheetName);
    if (localSheet && localSheet.getLastRow() > 1) {
      const result = searchSheetForJob(localSheet, cleanJob, sheetName);
      if (result && result.found) {
        return result;
      }
    }
  }

  // Tier 3: Published CSV Fallback (Used when user account lacks Drive permissions)
  for (const source of JOB_DB_CSV_FALLBACKS) {
    try {
      const response = UrlFetchApp.fetch(source.url, { muteHttpExceptions: true });
      if (response.getResponseCode() !== 200) continue;

      const rows = Utilities.parseCsv(response.getContentText());
      if (!rows || rows.length < 2) continue;

      const headers = rows[0].map(h => (h ? h.toString().trim().toUpperCase() : ""));
      const findHeaderIndex = (aliases, defaultIdx) => {
        for (const alias of aliases) {
          const idx = headers.indexOf(alias);
          if (idx !== -1) return idx;
        }
        return defaultIdx;
      };

      const jobCol = findHeaderIndex(["JOB", "JOB ORDER", "JOB ORDER NUMBER", "ORDER", "ORDER NO", "ORDER #", "JOB#", "CO#"], 7);
      const itemCol = findHeaderIndex(["ITEM", "ITEM NO", "ITEM NO.", "ITEM NUMBER", "ITEM#", "PART NUMBER", "PART NO", "PART"], 14);
      const coordCol = findHeaderIndex(["PROJECT COORDINATOR", "COORDINATOR", "PROJECT COORD", "PC", "PM"], 19);

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (row[jobCol] && row[jobCol].toString().trim().toUpperCase() === cleanJob) {
          return {
            found: true,
            itemNo: (itemCol !== -1 && row[itemCol]) ? row[itemCol].toString().trim() : "",
            projectCoordinator: (coordCol !== -1 && row[coordCol] && row[coordCol].toString().trim())
              ? row[coordCol].toString().trim()
              : (source.name === "STOCK ITEMS" ? "Stock Item" : "N/A"),
            sourceTab: source.name
          };
        }
      }
    } catch (csvErr) {
      console.warn(`CSV fallback failed for ${source.name}: ${csvErr}`);
    }
  }

  return { found: false };
}

/**
 * High-Speed Real-Time Sheet Search using TextFinder
 * Scans the Job column directly for sub-second lookup
 */
function searchSheetForJob(sheet, cleanJob, tabName) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return null;

  // Retrieve Row 1 headers to locate columns dynamically
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => (h ? h.toString().trim().toUpperCase() : ""));
  
  const findHeaderIndex = (aliases, defaultIdx) => {
    for (const alias of aliases) {
      const idx = headerRow.indexOf(alias);
      if (idx !== -1) return idx;
    }
    return defaultIdx;
  };

  const jobIdx = findHeaderIndex(["JOB", "JOB ORDER", "JOB ORDER NUMBER", "ORDER", "ORDER NO", "ORDER #", "JOB#", "CO#"], 7);
  const itemIdx = findHeaderIndex(["ITEM", "ITEM NO", "ITEM NO.", "ITEM NUMBER", "ITEM#", "PART NUMBER", "PART NO", "PART"], 14);
  const coordIdx = findHeaderIndex(["PROJECT COORDINATOR", "COORDINATOR", "PROJECT COORD", "PC", "PM"], 19);

  // Target the specific Job Order column with TextFinder for fast execution
  const searchRange = (jobIdx !== -1 && jobIdx < lastCol) 
    ? sheet.getRange(2, jobIdx + 1, lastRow - 1, 1) 
    : sheet.getDataRange();

  const match = searchRange.createTextFinder(cleanJob).matchEntireCell(true).findNext();
  if (!match) return null;

  // Read the matching row values
  const matchedRow = match.getRow();
  const rowValues = sheet.getRange(matchedRow, 1, 1, lastCol).getValues()[0];
  
  const itemNo = (itemIdx !== -1 && rowValues[itemIdx]) ? rowValues[itemIdx].toString().trim() : "";
  let coordinator = (coordIdx !== -1 && rowValues[coordIdx]) ? rowValues[coordIdx].toString().trim() : "";
  
  if (!coordinator && tabName === "STOCK ITEMS") {
    coordinator = "Stock Item";
  }

  return {
    found: true,
    itemNo: itemNo,
    projectCoordinator: coordinator || "N/A",
    sourceTab: tabName
  };
}

/**
 * Utility Function: Downloads the 3 published CSVs into local tabs
 * Populates tabs 'OOR', 'New Orders', and 'STOCK ITEMS'
 */
function syncJobDatabaseTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let syncCount = 0;

  for (const source of JOB_DB_CSV_FALLBACKS) {
    try {
      const response = UrlFetchApp.fetch(source.url, { muteHttpExceptions: true });
      if (response.getResponseCode() !== 200) {
        console.error(`Failed to fetch ${source.name} (HTTP ${response.getResponseCode()})`);
        continue;
      }

      const csvData = Utilities.parseCsv(response.getContentText());
      if (!csvData || csvData.length === 0) continue;

      let sheet = ss.getSheetByName(source.name);
      if (!sheet) {
        sheet = ss.insertSheet(source.name);
      } else {
        sheet.clearContents();
      }

      sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
      syncCount++;
    } catch (e) {
      console.error(`Sync error on tab ${source.name}: ${e}`);
    }
  }

  // Clear stale asset caches
  CacheService.getScriptCache().removeAll(['asset_ids']);
  ss.toast(`Job Database Sync: Refreshed ${syncCount} of ${JOB_DB_CSV_FALLBACKS.length} tabs.`, 'Sync Completed', 5);
}

/**
 * Processes form submission from BorrowDialog.html
 */
function processBorrowForm(formObject) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return "Error: System is currently busy with another operation. Please scan again.";
  }

  try {
    const projectCoordinator = (formObject.projectCoordinator || "N/A").trim();
    const assetId = (formObject.assetId || "").trim().toUpperCase();
    const jobOrder = (formObject.jobOrder || "N/A").trim();
    const itemNo = (formObject.itemNo || "").trim();
    
    if (!assetId) return "Error: Asset ID is required.";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const borrowSheet = ss.getSheetByName(BORROW_SHEET_NAME);
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!masterSheet) return `Error: Master sheet '${MASTER_SHEET_NAME}' not found.`;
    if (!borrowSheet) return `Error: Borrow sheet '${BORROW_SHEET_NAME}' not found.`;

    const masterData = masterSheet.getDataRange().getValues();
    if (masterData.length < 2) return "Error: Master sheet contains no asset data.";

    const colIndices = getColumnIndices(masterData[0]);
    if (colIndices.assetId === -1 || colIndices.status === -1) {
      return "Error: Could not find 'Asset' or 'Status' column in Master Sheet.";
    }

    // Verify Asset status in Borrow Log
    const borrowData = borrowSheet.getDataRange().getValues();
    for (let i = 1; i < borrowData.length; i++) {
      if (borrowData[i][3] && borrowData[i][3].toString().trim().toUpperCase() === assetId && borrowData[i][6] === "") { 
        return `Error: Asset ID '${assetId}' is already borrowed. Please return it first.`;
      }
    }

    // Locate Asset in Master Sheet
    let masterRowIndex = -1;
    let assetDescription = "";

    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][colIndices.assetId] && masterData[i][colIndices.assetId].toString().trim().toUpperCase() === assetId) {
        masterRowIndex = i + 1;
        assetDescription = colIndices.desc !== -1 ? masterData[i][colIndices.desc] : "";
        break;
      }
    }

    if (masterRowIndex === -1) return `Error: Asset ID '${assetId}' not found in the master list.`;

    // Update Master Sheet
    if (colIndices.assignedTo !== -1) {
      masterSheet.getRange(masterRowIndex, colIndices.assignedTo + 1).setValue(projectCoordinator);
    }
    masterSheet.getRange(masterRowIndex, colIndices.status + 1).setValue('Checked Out');

    // Add entry to Borrow Sheet in a single batched write
    borrowSheet.insertRowAfter(1);
    const newBorrowRow = [[jobOrder, itemNo, projectCoordinator, assetId, assetDescription, new Date(), ""]];
    borrowSheet.getRange(2, 1, 1, 7).setValues(newBorrowRow);

    return `Success: Asset '${assetId}' borrowed for Job '${jobOrder}'.`;

  } catch (e) {
    return "Error: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Processes form submission from ReturnDialog.html
 */
function processReturnForm(formObject) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return "Error: System busy. Please scan again.";
  }

  try {
    const assetId = (formObject.assetId || "").trim().toUpperCase();
    if (!assetId) return "Error: Asset ID is required.";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const borrowSheet = ss.getSheetByName(BORROW_SHEET_NAME);
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!masterSheet) return "Error: Master sheet not found.";
    if (!borrowSheet) return "Error: Borrow sheet not found.";
    
    const borrowData = borrowSheet.getDataRange().getValues();
    let borrowRowIndex = -1;

    for (let i = 1; i < borrowData.length; i++) {
      if (borrowData[i][3] && borrowData[i][3].toString().trim().toUpperCase() === assetId && borrowData[i][6] === '') {
        borrowRowIndex = i + 1;
        break; 
      }
    }

    if (borrowRowIndex === -1) return `Error: Asset ID '${assetId}' is not currently borrowed.`;

    // Record Return Timestamp
    borrowSheet.getRange(borrowRowIndex, 7).setValue(new Date()); 

    // Reset Master Sheet
    const masterData = masterSheet.getDataRange().getValues();
    const colIndices = getColumnIndices(masterData[0]);

    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][colIndices.assetId] && masterData[i][colIndices.assetId].toString().trim().toUpperCase() === assetId) {
        if (colIndices.assignedTo !== -1) {
          masterSheet.getRange(i + 1, colIndices.assignedTo + 1).setValue(''); 
        }
        if (colIndices.status !== -1) {
          masterSheet.getRange(i + 1, colIndices.status + 1).setValue('Available'); 
        }
        break;
      }
    }
    
    return `Success: Asset '${assetId}' has been returned.`;

  } catch (e) {
    return "Error: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Unified smart importer. Auto-detects file format (Standard Asset CSV vs SyteLine Stockroom Locations CSV).
 */
function importNewAssets(csvText) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return "Error: Import operation busy. Try again.";

  try {
    if (!csvText || !csvText.trim()) return "Error: Empty file uploaded.";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) return `Error: Master sheet '${MASTER_SHEET_NAME}' not found.`;

    const masterHeaders = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    const colIndices = getColumnIndices(masterHeaders);

    if (colIndices.assetId === -1) {
      return "Error: Master sheet must have an 'Asset' or 'Asset ID' column header in Row 1.";
    }

    const masterData = masterSheet.getDataRange().getValues();
    const existingAssetIds = new Set();
    
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][colIndices.assetId]) {
        existingAssetIds.add(masterData[i][colIndices.assetId].toString().trim().toUpperCase());
      }
    }

    // Robust delimiter detection by sampling the first line
    const firstLine = csvText.split(/\r\n|\n|\r/)[0] || '';
    const delimiter = (firstLine.match(/\t/g) || []).length > (firstLine.match(/,/g) || []).length ? '\t' : ',';
    
    const csvData = Utilities.parseCsv(csvText, delimiter);
    if (csvData.length < 2) return "Error: Uploaded file contains no data rows.";

    const fileHeaders = csvData[0];
    
    const findExactHeader = (names) => {
      const normalized = fileHeaders.map(h => h ? h.toString().trim().toUpperCase() : "");
      for (const name of names) {
        const idx = normalized.indexOf(name);
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const fileItemIdx = findExactHeader(["ITEM", "PART NUMBER", "PART NO", "PART"]);
    const fileAssetIdx = findExactHeader(["ASSET", "ASSET ID", "TAG", "ASSET#"]);
    const fileDescIdx = findExactHeader(["DESCRIPTION", "DESC"]);
    const fileExtDescIdx = findExactHeader(["EXT. DESCRIPTION", "EXTENDED DESCRIPTION", "LOCATION DESCRIPTION"]);
    const fileLocIdx = findExactHeader(["LOCATION", "STOCKROOM LOCATION", "LOC", "BIN", "SHELF"]);
    const fileMfgIdx = findExactHeader(["PRODUCT CLASS", "MANUFACTURER", "MFG", "BRAND", "CLASS"]);

    const isStockroomFile = (fileItemIdx !== -1 && fileAssetIdx === -1);
    let newAssetsAdded = 0;
    const rowsToAdd = [];

    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      if (!row || row.length === 0) continue;

      let partNum = (fileItemIdx !== -1 && row[fileItemIdx]) ? row[fileItemIdx].toString().trim().replace(/\s+/g, ' ') : "";
      let assetId = (fileAssetIdx !== -1 && row[fileAssetIdx]) ? row[fileAssetIdx].toString().trim().toUpperCase() : "";
      let desc = (fileDescIdx !== -1 && row[fileDescIdx]) ? row[fileDescIdx].toString().trim().replace(/\s+/g, ' ') : "";
      let extDesc = (fileExtDescIdx !== -1 && row[fileExtDescIdx]) ? row[fileExtDescIdx].toString().trim().replace(/\s+/g, ' ') : "";
      let loc = (fileLocIdx !== -1 && row[fileLocIdx]) ? row[fileLocIdx].toString().trim().replace(/\s+/g, ' ') : "";
      let mfg = (fileMfgIdx !== -1 && row[fileMfgIdx]) ? row[fileMfgIdx].toString().trim().replace(/\s+/g, ' ') : "";

      if (isStockroomFile || !assetId) {
        if (!partNum) continue;
        assetId = partNum.toUpperCase();
        if (!loc) loc = "Mezzanine";
      }

      if (!assetId) continue;

      if (!existingAssetIds.has(assetId)) {
        const newRow = new Array(masterHeaders.length).fill("");

        if (colIndices.partNum !== -1) newRow[colIndices.partNum] = partNum;
        if (colIndices.desc !== -1) newRow[colIndices.desc] = desc;
        if (colIndices.assetId !== -1) newRow[colIndices.assetId] = assetId;
        if (colIndices.extDesc !== -1) newRow[colIndices.extDesc] = extDesc;
        if (colIndices.location !== -1) newRow[colIndices.location] = loc;
        if (colIndices.assignedTo !== -1) newRow[colIndices.assignedTo] = "";
        if (colIndices.status !== -1) newRow[colIndices.status] = "Available";
        if (colIndices.category !== -1) newRow[colIndices.category] = isStockroomFile ? "RAW MATERIAL" : "EQUIPMENT";
        if (colIndices.group !== -1) newRow[colIndices.group] = isStockroomFile ? "REELS & SPOOLS" : "TOOLS";
        if (colIndices.mfg !== -1) newRow[colIndices.mfg] = mfg;

        rowsToAdd.push(newRow);
        newAssetsAdded++;
        existingAssetIds.add(assetId);
      }
    }

    if (rowsToAdd.length > 0) {
      masterSheet.getRange(masterSheet.getLastRow() + 1, 1, rowsToAdd.length, masterHeaders.length).setValues(rowsToAdd);
      CacheService.getScriptCache().remove('asset_ids');
    }

    return `Import complete: ${newAssetsAdded} new records added to Master List.`;

  } catch (e) {
    return "Error: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Searches Master Sheet & Borrow Log for asset details (Strictly Read-Only)
 */
function findAsset(formObject) {
  try {
    const assetId = (formObject.assetId || "").toString().trim().toUpperCase();
    if (!assetId) return "Error: Please enter or scan an Asset ID.";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) return "Error: Master sheet not found.";

    const masterData = masterSheet.getDataRange().getValues();
    if (masterData.length < 2) return `Error: Asset ID '${assetId}' not found.`;

    const colIndices = getColumnIndices(masterData[0]);

    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][colIndices.assetId] && masterData[i][colIndices.assetId].toString().trim().toUpperCase() === assetId) {
        let currentStatus = colIndices.status !== -1 ? masterData[i][colIndices.status] : "Unknown";
        let assignedTo = colIndices.assignedTo !== -1 ? masterData[i][colIndices.assignedTo] : "";
        const description = colIndices.desc !== -1 ? masterData[i][colIndices.desc] : "";
        const location = colIndices.location !== -1 ? masterData[i][colIndices.location] : "";
        
        let jobOrder = "";
        let itemNo = "";
        let borrowDateStr = "";
        
        const borrowSheet = ss.getSheetByName(BORROW_SHEET_NAME);
        if (borrowSheet) {
          const borrowData = borrowSheet.getDataRange().getValues();
          for (let j = 1; j < borrowData.length; j++) {
            if (borrowData[j][3] && borrowData[j][3].toString().trim().toUpperCase() === assetId && borrowData[j][6] === "") {
              currentStatus = "Checked Out";
              jobOrder = borrowData[j][0] || "N/A"; 
              itemNo = borrowData[j][1] || "N/A";
              assignedTo = borrowData[j][2] || assignedTo || "N/A"; 
              if (borrowData[j][5]) borrowDateStr = new Date(borrowData[j][5]).toLocaleDateString();
              break;
            }
          }
        }
        
        let message = `Asset ID: ${assetId}\nDescription: ${description}\nLocation: ${location || 'N/A'}\nStatus: ${currentStatus}`;
        if (currentStatus === 'Checked Out') {
           if (jobOrder) message += `\nJob Order: ${jobOrder}`;
           if (itemNo) message += `\nItem No.: ${itemNo}`;
           if (assignedTo) message += `\nProject Coordinator: ${assignedTo}`;
           if (borrowDateStr) message += `\nBorrowed On: ${borrowDateStr}`;
        }
        return message;
      }
    }
    return `Error: Asset ID '${assetId}' not found.`;
  } catch (e) {
    return "Error: " + e.toString();
  }
}

/**
 * Returns array of Asset IDs for UI autocomplete
 */
function getAssetIds() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('asset_ids');
  if (cached) return JSON.parse(cached);

  try {
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) return [];

    const masterData = masterSheet.getDataRange().getValues();
    if (masterData.length < 2) return [];

    const colIndices = getColumnIndices(masterData[0]);
    if (colIndices.assetId === -1) return [];

    const assetIds = [];
    for (let i = 1; i < masterData.length; i++) {
      const val = masterData[i][colIndices.assetId];
      if (val) assetIds.push(val.toString().trim());
    }

    cache.put('asset_ids', JSON.stringify(assetIds), 600);
    return assetIds;
  } catch (e) {
    return [];
  }
}
