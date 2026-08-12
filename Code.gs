// Global constants for internal sheet names
const BORROW_SHEET_NAME = "Borrow Tools";
const MASTER_SHEET_NAME = "ToExcel_MTL_AssetManagementTable";

// *** CONFIGURATION FOR EXTERNAL JOB DB ***
const EXTERNAL_JOB_DB_ID = '1vGPJvUOgGu7xEehsXu82QFM04qdo513pW8r3XFnzJRM'; 
// Define multiple sheets to search through. Add or remove sheet names as necessary.
const EXTERNAL_JOB_DB_SHEET_NAMES = ['OOR', 'New Orders', '2026 WK1 to WK52']; 

/**
 * onOpen hook to build custom Google Sheets menu
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
 * Dialog Display Functions
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
 * Helper function to map header column names to 0-based indices dynamically.
 * Fallbacks to default index if column name is not found.
 */
function getColumnIndices(sheet, defaultMap) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  
  // Normalize header names to uppercase stripped strings
  headers.forEach((h, index) => {
    if (h) {
      const cleanHeader = h.toString().trim().toUpperCase();
      map[cleanHeader] = index;
    }
  });

  const resolvedIndices = {};
  for (const [key, defaultIdx] of Object.entries(defaultMap)) {
    let foundIdx = -1;
    for (const [headerName, idx] of Object.entries(map)) {
      if (headerName === key || headerName.includes(key) || key.includes(headerName)) {
        foundIdx = idx;
        break;
      }
    }
    resolvedIndices[key] = foundIdx !== -1 ? foundIdx : defaultIdx;
  }
  return resolvedIndices;
}

/**
 * Fetches Item No and Project Coordinator from external sheets based on Job Order
 */
function getJobDetails(jobOrder) {
  if (!jobOrder) return null;
  const cleanJobOrder = jobOrder.toString().trim().toUpperCase();

  // Check cache first to avoid slow external sheet opens
  const cache = CacheService.getScriptCache();
  const cacheKey = 'job_' + cleanJobOrder;
  const cachedData = cache.get(cacheKey);
  if (cachedData) {
    try {
      return JSON.parse(cachedData);
    } catch (e) {
      // Continue to fetch if cache parse fails
    }
  }

  try {
    const ss = SpreadsheetApp.openById(EXTERNAL_JOB_DB_ID);
    
    // Iterate through designated external sheets
    for (const sheetName of EXTERNAL_JOB_DB_SHEET_NAMES) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) continue;

      // Find Job Order column dynamically (default index 7)
      let jobColIdx = 7;
      let itemNoColIdx = 14;
      let pcColIdx = 19;

      const headers = data[0];
      headers.forEach((h, idx) => {
        const str = h.toString().trim().toUpperCase();
        if (str.includes("JOB ORDER") || str === "JOB") jobColIdx = idx;
        if (str.includes("ITEM NO") || str.includes("ITEM")) itemNoColIdx = idx;
        if (str.includes("PROJECT COORDINATOR") || str.includes("PC")) pcColIdx = idx;
      });

      for (let i = 1; i < data.length; i++) {
        if (data[i][jobColIdx] && data[i][jobColIdx].toString().trim().toUpperCase() === cleanJobOrder) {
          const result = {
            found: true,
            itemNo: data[i][itemNoColIdx] || "N/A",
            projectCoordinator: data[i][pcColIdx] || "N/A"
          };
          // Cache successful lookup for 30 minutes (1800 seconds)
          cache.put(cacheKey, JSON.stringify(result), 1800);
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
 * Processes the form submission from BorrowDialog.html.
 */
function processBorrowForm(formObject) {
  const lock = LockService.getScriptLock();
  // Wait up to 10 seconds for concurrent locks to clear
  if (!lock.tryLock(10000)) {
    return "Error: System is currently busy processing another request. Please try again.";
  }

  try {
    const projectCoordinator = formObject.projectCoordinator || "N/A";
    const pcName = projectCoordinator; 
    
    const assetId = (formObject.assetId || "").toString().trim().toUpperCase();
    const jobOrder = formObject.jobOrder || "N/A";
    const itemNo = formObject.itemNo || "N/A";
    
    if (!assetId) return "Error: Asset ID is required.";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const borrowSheet = ss.getSheetByName(BORROW_SHEET_NAME);
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!masterSheet) return "Error: Master asset sheet '" + MASTER_SHEET_NAME + "' not found.";
    if (!borrowSheet) return "Error: Borrow sheet '" + BORROW_SHEET_NAME + "' not found.";

    // Dynamic column resolution for Borrow sheet
    const borrowCols = getColumnIndices(borrowSheet, {
      "JOB ORDER": 0,
      "ITEM NO": 1,
      "PC": 2,
      "ASSET ID": 3,
      "DESCRIPTION": 4,
      "BORROW DATE": 5,
      "RETURN DATE": 6
    });

    const borrowData = borrowSheet.getDataRange().getValues();
    for (let i = 1; i < borrowData.length; i++) {
      const row = borrowData[i];
      const rowAssetId = (row[borrowCols["ASSET ID"]] || "").toString().trim().toUpperCase();
      const returnDate = row[borrowCols["RETURN DATE"]];
      
      if (rowAssetId === assetId && (returnDate === "" || returnDate === null || returnDate === undefined)) { 
        return `Error: Asset ID '${assetId}' is already borrowed. Please return it first.`;
      }
    }

    // Dynamic column resolution for Master sheet
    const masterCols = getColumnIndices(masterSheet, {
      "ASSET": 2,
      "DESCRIPTION": 3,
      "ASSIGNED TO": 5,
      "STATUS": 9
    });

    const masterData = masterSheet.getDataRange().getValues();
    let assetFound = false;
    let assetDescription = "";
    let masterRowIndex = -1;

    for (let i = 1; i < masterData.length; i++) {
      const currentId = (masterData[i][masterCols["ASSET"]] || "").toString().trim().toUpperCase();
      if (currentId === assetId) {
        assetFound = true;
        assetDescription = masterData[i][masterCols["DESCRIPTION"]] || "";
        masterRowIndex = i + 1; // 1-based sheet row index
        break;
      }
    }

    if (!assetFound) return `Error: Asset ID '${assetId}' not found in master inventory.`;

    // Update Master Sheet (1-based column indices)
    masterSheet.getRange(masterRowIndex, masterCols["ASSIGNED TO"] + 1).setValue(pcName); 
    masterSheet.getRange(masterRowIndex, masterCols["STATUS"] + 1).setValue('Checked Out');

    // Add entry to Borrow Sheet in a single batched array operation
    borrowSheet.insertRowAfter(1);
    const newBorrowRow = [];
    newBorrowRow[borrowCols["JOB ORDER"]] = jobOrder;
    newBorrowRow[borrowCols["ITEM NO"]] = itemNo;
    newBorrowRow[borrowCols["PC"]] = projectCoordinator;
    newBorrowRow[borrowCols["ASSET ID"]] = assetId;
    newBorrowRow[borrowCols["DESCRIPTION"]] = assetDescription;
    newBorrowRow[borrowCols["BORROW DATE"]] = new Date();
    newBorrowRow[borrowCols["RETURN DATE"]] = "";

    // Fill array gaps if any
    for (let c = 0; c < Math.max(...Object.values(borrowCols)) + 1; c++) {
      if (newBorrowRow[c] === undefined) newBorrowRow[c] = "";
    }

    borrowSheet.getRange(2, 1, 1, newBorrowRow.length).setValues([newBorrowRow]);

    return `Success: Asset '${assetId}' borrowed for Job '${jobOrder}'.`;

  } catch (e) {
    return "Error: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Processes the form submission from ReturnDialog.html.
 */
function processReturnForm(formObject) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return "Error: System is currently busy processing another request. Please try again.";
  }

  try {
    const assetId = (formObject.assetId || "").toString().trim().toUpperCase();
    if (!assetId) return "Error: Asset ID is required.";
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const borrowSheet = ss.getSheetByName(BORROW_SHEET_NAME);
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!masterSheet) return "Error: Master sheet not found.";
    if (!borrowSheet) return "Error: Borrow sheet not found.";
    
    const borrowCols = getColumnIndices(borrowSheet, {
      "ASSET ID": 3,
      "RETURN DATE": 6
    });

    const borrowData = borrowSheet.getDataRange().getValues();
    let borrowRowIndex = -1;

    // Find latest open borrow record
    for (let i = 1; i < borrowData.length; i++) {
      const rowAssetId = (borrowData[i][borrowCols["ASSET ID"]] || "").toString().trim().toUpperCase();
      const returnDate = borrowData[i][borrowCols["RETURN DATE"]];
      
      if (rowAssetId === assetId && (returnDate === '' || returnDate === null || returnDate === undefined)) {
        borrowRowIndex = i + 1;
        break; 
      }
    }

    if (borrowRowIndex === -1) return `Error: Asset ID '${assetId}' is not currently marked as borrowed.`;

    // Update Borrow sheet return date
    borrowSheet.getRange(borrowRowIndex, borrowCols["RETURN DATE"] + 1).setValue(new Date()); 

    // Update Master sheet status
    const masterCols = getColumnIndices(masterSheet, {
      "ASSET": 2,
      "ASSIGNED TO": 5,
      "STATUS": 9
    });

    const masterData = masterSheet.getDataRange().getValues();
    let masterRowIndex = -1;
    for (let i = 1; i < masterData.length; i++) {
      const currentId = (masterData[i][masterCols["ASSET"]] || "").toString().trim().toUpperCase();
      if (currentId === assetId) {
        masterRowIndex = i + 1;
        break;
      }
    }

    if (masterRowIndex !== -1) {
      masterSheet.getRange(masterRowIndex, masterCols["ASSIGNED TO"] + 1).setValue(''); 
      masterSheet.getRange(masterRowIndex, masterCols["STATUS"] + 1).setValue('Available'); 
    }
    
    return `Success: Asset '${assetId}' has been returned.`;

  } catch (e) {
    return "Error: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Imports new assets from standard CSV or TSV data.
 * Auto-detects delimiters and clears autocomplete cache.
 */
function importNewAssets(csvText) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return "Error: Import system busy. Please try again in a few seconds.";
  }

  try {
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) return "Error: Master sheet not found.";

    const masterCols = getColumnIndices(masterSheet, {
      "ASSET": 2,
      "ASSIGNED TO": 5,
      "STATUS": 9
    });

    const existingAssetIds = new Set(
      masterSheet.getRange(2, masterCols["ASSET"] + 1, Math.max(1, masterSheet.getLastRow() - 1), 1)
        .getValues()
        .flat()
        .map(id => id ? id.toString().trim().toUpperCase() : "")
        .filter(Boolean)
    );

    // Auto-detect CSV vs TSV delimiter
    const delimiter = csvText.includes('\t') ? '\t' : ',';
    const csvData = Utilities.parseCsv(csvText, delimiter);
    let newAssetsAdded = 0;
    const rowsToAdd = [];

    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      if (!row || row.length === 0) continue;
      
      const csvAssetId = (row[masterCols["ASSET"]] || "").toString().trim().toUpperCase();
      if (!csvAssetId) continue;

      if (!existingAssetIds.has(csvAssetId)) {
        row[masterCols["STATUS"]] = "Available"; 
        row[masterCols["ASSIGNED TO"]] = "";
        rowsToAdd.push(row);
        newAssetsAdded++;
        existingAssetIds.add(csvAssetId);
      }
    }

    if (rowsToAdd.length > 0) {
      masterSheet.getRange(masterSheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
      
      // Invalidate asset ID cache so autocomplete updates immediately
      CacheService.getScriptCache().remove('asset_ids');
    }

    return `Import complete: ${newAssetsAdded} new assets were added.`;

  } catch (e) {
    return "Error: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Find Logic - Pure read-only lookup of asset status and borrower details.
 */
function findAsset(formObject) {
  try {
    const assetId = (formObject.assetId || "").toString().trim().toUpperCase();
    if (!assetId) return "Error: Asset ID is required.";

    const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) return "Error: Master sheet not found.";

    const masterCols = getColumnIndices(masterSheet, {
      "ASSET": 2,
      "DESCRIPTION": 3,
      "ASSIGNED TO": 5,
      "STATUS": 9
    });

    const data = masterSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const currentId = (data[i][masterCols["ASSET"]] || "").toString().trim().toUpperCase();
      
      if (currentId === assetId) {
        let currentStatus = data[i][masterCols["STATUS"]] || "Available";
        let assignedTo = data[i][masterCols["ASSIGNED TO"]] || "";
        const description = data[i][masterCols["DESCRIPTION"]] || "N/A";
        
        let realStatus = currentStatus;
        let jobOrder = "";
        let itemNo = "";
        let borrowDateStr = "";
        
        const borrowSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BORROW_SHEET_NAME);
        if (borrowSheet) {
          const borrowCols = getColumnIndices(borrowSheet, {
            "JOB ORDER": 0,
            "ITEM NO": 1,
            "PC": 2,
            "ASSET ID": 3,
            "BORROW DATE": 5,
            "RETURN DATE": 6
          });

          const borrowData = borrowSheet.getDataRange().getValues();
          for (let j = 1; j < borrowData.length; j++) {
            const borrowRow = borrowData[j];
            const borrowAssetId = (borrowRow[borrowCols["ASSET ID"]] || "").toString().trim().toUpperCase();
            
            if (borrowAssetId === assetId) {
              if (borrowRow[borrowCols["BORROW DATE"]]) {
                borrowDateStr = new Date(borrowRow[borrowCols["BORROW DATE"]]).toLocaleDateString();
              }

              const returnDate = borrowRow[borrowCols["RETURN DATE"]];
              if (returnDate === "" || returnDate === null || returnDate === undefined) {
                realStatus = "Checked Out";
                jobOrder = borrowRow[borrowCols["JOB ORDER"]] ? borrowRow[borrowCols["JOB ORDER"]] : "N/A"; 
                itemNo = borrowRow[borrowCols["ITEM NO"]] ? borrowRow[borrowCols["ITEM NO"]] : "N/A";
                assignedTo = borrowRow[borrowCols["PC"]] ? borrowRow[borrowCols["PC"]] : assignedTo; 
              } else {
                realStatus = "Available";
              }
              break;
            }
          }
        }
        
        let message = `Asset ID: ${assetId}\nDescription: ${description}\nStatus: ${realStatus}`;
        if (realStatus === 'Checked Out') {
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
 * Get Asset IDs for autocomplete with robust caching
 */
function getAssetIds() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('asset_ids');
  if (cached != null) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  try {
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) return [];

    const masterCols = getColumnIndices(masterSheet, { "ASSET": 2 });
    const colIndex = masterCols["ASSET"] + 1; // 1-based column

    const lastRow = masterSheet.getLastRow();
    if (lastRow < 2) return [];

    const data = masterSheet.getRange(2, colIndex, lastRow - 1, 1)
      .getValues()
      .flat()
      .map(id => id ? id.toString().trim() : "")
      .filter(Boolean);

    // Deduplicate
    const uniqueData = Array.from(new Set(data));

    cache.put('asset_ids', JSON.stringify(uniqueData), 600); // 10 minutes cache
    return uniqueData;
  } catch (e) {
    return [];
  }
}
