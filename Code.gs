// Global constants for internal sheet names
const BORROW_SHEET_NAME = "Borrow Tools";
const MASTER_SHEET_NAME = "ToExcel_MTL_AssetManagementTable";

// *** CONFIGURATION FOR EXTERNAL JOB DB ***
const EXTERNAL_JOB_DB_ID = '1vGPJvUOgGu7xEehsXu82QFM04qdo513pW8r3XFnzJRM'; 
const EXTERNAL_JOB_DB_SHEET_NAME = 'OOR'; 

/**
 * onOpen
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
 * Fetches Item No and Project Coordinator from external sheet based on Job Order
 */
function getJobDetails(jobOrder) {
  if (!jobOrder) return null;
  
  try {
    const ss = SpreadsheetApp.openById(EXTERNAL_JOB_DB_ID);
    const sheet = ss.getSheetByName(EXTERNAL_JOB_DB_SHEET_NAME);
    if (!sheet) return { error: "External sheet '" + EXTERNAL_JOB_DB_SHEET_NAME + "' not found. Check the tab name." };
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      // Check Index 7 (Job Order)
      if (data[i][7] && data[i][7].toString().toUpperCase() === jobOrder.toUpperCase()) {
        return {
          found: true,
          itemNo: data[i][14],           // Index 14 (Item No.)
          projectCoordinator: data[i][19] // Index 19 (Project Coordinator)
        };
      }
    }
    
    return { found: false };
    
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * Processes the form submission from the BorrowDialog.html.
 */
function processBorrowForm(formObject) {
  try {
    const projectCoordinator = formObject.projectCoordinator || "N/A";
    const pcName = projectCoordinator; 
    
    const assetId = formObject.assetId.toUpperCase();
    const jobOrder = formObject.jobOrder || "N/A";
    const itemNo = formObject.itemNo || "";
    
    const borrowSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BORROW_SHEET_NAME);
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);

    if (!masterSheet) return "Error: Master asset sheet '" + MASTER_SHEET_NAME + "' not found.";
    if (!borrowSheet) return "Error: Borrow sheet '" + BORROW_SHEET_NAME + "' not found.";

    // Logic: Check if the tool is already borrowed and not returned
    const borrowData = borrowSheet.getDataRange().getValues();
    for (let i = 1; i < borrowData.length; i++) {
      const row = borrowData[i];
      // Check Col D (Index 3) for ID and Col G (Index 6) for Return Date
      if (row[3].toString().toUpperCase() === assetId && row[6] === "") { 
        // DISABLED REDIRECT:
        // showReturnDialog(assetId); 
        return `Error: Asset ID '${assetId}' is already borrowed. Please return it first.`;
      }
    }

    // Find Asset in Master List
    const masterData = masterSheet.getDataRange().getValues();
    let assetFound = false;
    let assetDescription = "";
    let rowIndex = -1;

    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][2].toString().toUpperCase() === assetId) {
        assetFound = true;
        assetDescription = masterData[i][3];
        rowIndex = i + 1;
        break;
      }
    }

    if (!assetFound) return `Error: Asset ID '${assetId}' not found in the master list.`;

    // Update Master Sheet
    masterSheet.getRange(rowIndex, 6).setValue(pcName); 
    masterSheet.getRange(rowIndex, 10).setValue('Checked Out');

    // Add entry to Borrow Sheet
    borrowSheet.insertRowAfter(1);
    borrowSheet.getRange('A2').setValue(jobOrder);           // Col A: Job Order
    borrowSheet.getRange('B2').setValue(itemNo);             // Col B: Item No
    borrowSheet.getRange('C2').setValue(projectCoordinator); // Col C: Project Coordinator
    borrowSheet.getRange('D2').setValue(assetId);            // Col D: Asset ID
    borrowSheet.getRange('E2').setValue(assetDescription);   // Col E: Description
    borrowSheet.getRange('F2').setValue(new Date());         // Col F: Borrow Date
    // Col G: Return Date (Left blank)

    return `Success: Asset '${assetId}' borrowed for Job '${jobOrder}'.`;

  } catch (e) {
    return "Error: " + e.toString();
  }
}

/**
 * Processes the form submission from the ReturnDialog.html.
 */
function processReturnForm(formObject) {
  try {
    const assetId = formObject.assetId.toUpperCase();
    
    const borrowSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BORROW_SHEET_NAME);
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);

    if (!masterSheet) return "Error: Master sheet not found.";
    if (!borrowSheet) return "Error: Borrow sheet not found.";
    
    const borrowData = borrowSheet.getDataRange().getValues();
    let borrowRowIndex = -1;

    // Find latest open borrow record
    for (let i = 1; i < borrowData.length; i++) {
      if (borrowData[i][3].toString().toUpperCase() === assetId && borrowData[i][6] === '') {
        borrowRowIndex = i + 1;
        break; 
      }
    }

    if (borrowRowIndex === -1) return `Error: Asset ID '${assetId}' is not currently borrowed.`;

    // Update Borrow sheet
    borrowSheet.getRange(borrowRowIndex, 7).setValue(new Date()); 

    // Update Master sheet
    const masterData = masterSheet.getDataRange().getValues();
    let masterRowIndex = -1;
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][2].toString().toUpperCase() === assetId) {
        masterRowIndex = i + 1;
        break;
      }
    }

    if (masterRowIndex !== -1) {
      masterSheet.getRange(masterRowIndex, 6).setValue(''); 
      masterSheet.getRange(masterRowIndex, 10).setValue('Available'); 
    }
    
    return `Success: Asset '${assetId}' has been returned.`;

  } catch (e) {
    return "Error: " + e.toString();
  }
}

/**
 * Import Logic
 */
function importNewAssets(csvText) {
  try {
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) return "Error: Master sheet not found.";

    const existingAssetIds = new Set(
      masterSheet.getRange(2, 3, masterSheet.getLastRow() - 1, 1).getValues().flat().map(id => id.toString().toUpperCase())
    );

    const csvData = Utilities.parseCsv(csvText, '\t');
    let newAssetsAdded = 0;
    const rowsToAdd = [];

    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      if (!row || !row[2]) continue;
      
      const csvAssetId = row[2].toString().toUpperCase();

      if (!existingAssetIds.has(csvAssetId)) {
        row[9] = "Available"; 
        row[5] = "";
        rowsToAdd.push(row);
        newAssetsAdded++;
        existingAssetIds.add(csvAssetId);
      }
    }

    if (rowsToAdd.length > 0) {
      masterSheet.getRange(masterSheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    }

    return `Import complete: ${newAssetsAdded} new assets were added.`;

  } catch (e) {
    return "Error: " + e.toString();
  }
}

/**
 * Find Logic
 */
function findAsset(formObject) {
  try {
    const assetId = formObject.assetId.toUpperCase();
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) return "Error: Master sheet not found.";

    const data = masterSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2].toString().toUpperCase() === assetId) {
        let currentStatus = data[i][9];
        let assignedTo = data[i][5];
        const description = data[i][3];
        
        let realStatus = currentStatus;
        let jobOrder = "";
        let itemNo = "";
        let borrowDateStr = "";
        
        const borrowSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BORROW_SHEET_NAME);
        if (borrowSheet) {
          const borrowData = borrowSheet.getDataRange().getValues();
          for (let j = 1; j < borrowData.length; j++) {
            const borrowRow = borrowData[j];
            
            // Match Asset ID (Col D, Index 3)
            if (borrowRow[3].toString().toUpperCase() === assetId) {
              
              if (borrowRow[5]) borrowDateStr = new Date(borrowRow[5]).toLocaleDateString();

              // Check Index 6 (Col G) for Return Date
              if (borrowRow[6] === "") {
                realStatus = "Checked Out";
                jobOrder = borrowRow[0] ? borrowRow[0] : "N/A"; 
                itemNo = borrowRow[1] ? borrowRow[1] : "N/A";
                assignedTo = borrowRow[2] ? borrowRow[2] : "N/A"; 
              } else {
                realStatus = "Available";
                if (currentStatus === "PENDING?") {
                   masterSheet.getRange(i+1, 10).setValue("Available");
                }
              }
              break;
            }
          }
        }
        
        if (realStatus !== currentStatus && currentStatus !== "PENDING?") {
           masterSheet.getRange(i + 1, 10).setValue(realStatus);
           masterSheet.getRange(i + 1, 6).setValue(realStatus === 'Available' ? '' : assignedTo);
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
 * Get Asset IDs for autocomplete
 */
function getAssetIds() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('asset_ids');
  if (cached != null) return JSON.parse(cached);

  try {
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) return [];
    const data = masterSheet.getRange("C2:C").getValues().flat().filter(String);
    cache.put('asset_ids', JSON.stringify(data), 600);
    return data;
  } catch (e) {
    return [];
  }
}
