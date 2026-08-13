// Global constants for internal sheet names
const BORROW_SHEET_NAME = "Borrow Tools";
const MASTER_SHEET_NAME = "ToExcel_MTL_AssetManagementTable";

// *** CONFIGURATION FOR EXTERNAL JOB DB ***
const EXTERNAL_JOB_DB_ID = '1vGPJvUOgGu7xEehsXu82QFM04qdo513pW8r3XFnzJRM'; 
const EXTERNAL_JOB_DB_SHEET_NAMES = ['OOR', 'New Orders', '2026 WK1 to WK52']; 

/**
 * Creates custom spreadsheet menu on document open
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Asset Management')
      .addItem('Open Main Menu', 'showMainMenuDialog')
      .addSeparator()
      .addItem('Import New Assets', 'showImportDialog')
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
 * Dynamic Column Index Mapping Helper
 * Finds zero-based column indices by checking Row 1 header names against multiple aliases
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
 * Fetches Item No and Project Coordinator from external sheets based on Job Order
 */
function getJobDetails(jobOrder) {
  if (!jobOrder) return null;
  const cleanJob = jobOrder.toString().trim().toUpperCase();

  const cache = CacheService.getScriptCache();
  const cached = cache.get(`job_${cleanJob}`);
  if (cached) return JSON.parse(cached);

  try {
    const ss = SpreadsheetApp.openById(EXTERNAL_JOB_DB_ID);
    
    for (const sheetName of EXTERNAL_JOB_DB_SHEET_NAMES) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;

      for (let i = 1; i < data.length; i++) {
        if (data[i][7] && data[i][7].toString().trim().toUpperCase() === cleanJob) {
          const result = {
            found: true,
            itemNo: data[i][14] ? data[i][14].toString().trim() : "",
            projectCoordinator: data[i][19] ? data[i][19].toString().trim() : ""
          };
          cache.put(`job_${cleanJob}`, JSON.stringify(result), 300);
          return result;
        }
      }
    }
    return { found: false };
  } catch (e) {
    return { error: e.toString() };
  }
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

    // Add entry to Borrow Sheet in a single batched array write
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
 * Dynamically builds rows matching whatever columns exist in the Master Sheet.
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

    // Auto-detect delimiter (Tab vs Comma)
    const delimiter = csvText.includes('\t') ? '\t' : ',';
    const csvData = Utilities.parseCsv(csvText, delimiter);
    if (csvData.length < 2) return "Error: Uploaded file contains no data rows.";

    const fileHeaders = csvData[0];
    
    // Strict Header Search Helper
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

      // Use Item (Part Number) directly as the Asset ID for Stockroom imports
      if (isStockroomFile || !assetId) {
        if (!partNum) continue;
        assetId = partNum.toUpperCase();
        if (!loc) loc = "Mezzanine";
      }

      if (!assetId) continue;

      if (!existingAssetIds.has(assetId)) {
        // Construct clean row matching Master Sheet's exact column length
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
