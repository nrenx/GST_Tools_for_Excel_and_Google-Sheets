/**
 * ===============================================================================
 * MODULE: Module_Master
 * DESCRIPTION: Handles all operations related to the 'Master' sheet, including
 *              invoice record management and the invoice numbering system.
 * ===============================================================================
 */

// ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
// 📋 MASTER SHEET & INVOICE COUNTER FUNCTIONS
// ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

function CreateMasterSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Delete existing Master sheet if it exists
    const existingSheet = spreadsheet.getSheetByName("Master");
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    // Create new Master sheet
    const sheet = spreadsheet.insertSheet("Master");

    // ===== GST INVOICE RECORDS FOR AUDIT & RETURN FILING (A1:P1) =====
    // GST-compliant headers for complete invoice records
    const headers = [
      "Invoice_Number", "Invoice_Date", "Customer_Name", "Customer_GSTIN", "Customer_State",
      "Customer_State_Code", "Total_Taxable_Value", "IGST_Rate", "IGST_Amount", "Total_Tax_Amount",
      "Total_Invoice_Value", "HSN_Codes", "Item_Description", "Quantity", "UOM", "Date_Created"
    ];

    // Set headers
    for (let i = 0; i < headers.length; i++) {
      sheet.getRange(1, i + 1).setValue(headers[i]);
    }

    // Format GST audit headers
    const headerRange = sheet.getRange("A1:P1");
    headerRange.setFontWeight("bold")
               .setBackground("#2F5061")
               .setFontColor("#FFFFFF")
               .setHorizontalAlignment("center")
               .setWrap(true)
               .setBorder(true, true, true, true, false, false, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

    sheet.setRowHeight(1, 30);

    // Set specific column widths for GST data
    const columnWidths = [
      { col: 1, width: 160 },  // A: Invoice Number
      { col: 2, width: 120 },  // B: Invoice Date
      { col: 3, width: 240 },  // C: Customer Name
      { col: 4, width: 160 },  // D: Customer GSTIN
      { col: 5, width: 200 },  // E: Customer State
      { col: 6, width: 120 },  // F: State Code
      { col: 7, width: 160 },  // G: Taxable Value
      { col: 8, width: 96 },   // H: IGST Rate
      { col: 9, width: 120 },  // I: IGST Amount
      { col: 10, width: 120 }, // J: Total Tax
      { col: 11, width: 160 }, // K: Invoice Value
      { col: 12, width: 160 }, // L: HSN Codes
      { col: 13, width: 320 }, // M: Item Description
      { col: 14, width: 96 },  // N: Quantity
      { col: 15, width: 80 },  // O: UOM
      { col: 16, width: 160 }  // P: Date Created
    ];

    columnWidths.forEach(item => {
      sheet.setColumnWidth(item.col, item.width);
    });

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error creating Master sheet: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function GetNextInvoiceNumber() {
  try {
    // Ensure supporting worksheets exist
    EnsureAllSupportingWorksheetsExist();

    // Get or create Master sheet
    const masterSheet = GetOrCreateWorksheet("Master");
    
    const currentYear = new Date().getFullYear();
    let maxCounter = 0;

    // Find the highest counter for the current year by examining existing invoice records
    const lastRow = masterSheet.getLastRow();

    if (lastRow > 1) { // If there are invoice records
      const invoiceNumbers = masterSheet.getRange("A2:A" + lastRow).getValues();
      
      for (let i = 0; i < invoiceNumbers.length; i++) {
        const invoiceNum = invoiceNumbers[i][0].toString().trim();
        if (invoiceNum && invoiceNum.indexOf(`INV-${currentYear}-`) === 0) {
          // Extract counter from invoice number (format: INV-YYYY-NNN)
          const counterStr = invoiceNum.substring(invoiceNum.lastIndexOf('-') + 1);
          const counter = parseInt(counterStr, 10);
          if (!isNaN(counter)) {
            maxCounter = Math.max(maxCounter, counter);
          }
        }
      }
    }

    // Set next counter
    const nextCounter = maxCounter + 1;

    // Generate new invoice number
    const newInvoiceNumber = `INV-${currentYear}-${nextCounter.toString().padStart(3, '0')}`;

    return newInvoiceNumber;
  } catch (error) {
    console.error('Error getting next invoice number:', error);
    return `INV-${new Date().getFullYear()}-001`;
  }
}

function GetCurrentInvoiceNumber() {
  try {
    // Ensure supporting worksheets exist
    EnsureAllSupportingWorksheetsExist();

    const masterSheet = GetOrCreateWorksheet("Master");
    const currentYear = new Date().getFullYear();
    let maxCounter = 0;

    if (!masterSheet) {
      return `INV-${currentYear}-001`;
    }

    const lastRow = masterSheet.getLastRow();

    if (lastRow > 1) { // If there are invoice records
      const invoiceNumbers = masterSheet.getRange("A2:A" + lastRow).getValues();
      
      for (let i = 0; i < invoiceNumbers.length; i++) {
        const invoiceNum = invoiceNumbers[i][0].toString().trim();
        if (invoiceNum && invoiceNum.indexOf(`INV-${currentYear}-`) === 0) {
          // Extract counter from invoice number (format: INV-YYYY-NNN)
          const counterStr = invoiceNum.substring(invoiceNum.lastIndexOf('-') + 1);
          const counter = parseInt(counterStr, 10);
          if (!isNaN(counter)) {
            maxCounter = Math.max(maxCounter, counter);
          }
        }
      }
    }

    if (maxCounter === 0) {
      return `INV-${currentYear}-001`;
    }

    // Return current invoice number
    return `INV-${currentYear}-${maxCounter.toString().padStart(3, '0')}`;
  } catch (error) {
    console.error('Error getting current invoice number:', error);
    return `INV-${new Date().getFullYear()}-001`;
  }
}

function ResetInvoiceCounter() {
  // Reset invoice counter by clearing all records from Master sheet
  try {
    const response = SpreadsheetApp.getUi().alert(
      'Confirm Reset',
      'This will delete ALL invoice records from the Master sheet.\n' +
      'This action cannot be undone.\n\n' +
      'Are you sure you want to proceed?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (response !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }

    const masterSheet = GetOrCreateWorksheet("Master");
    
    if (!masterSheet) {
      SpreadsheetApp.getUi().alert('Error', 'Master sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const lastRow = masterSheet.getLastRow();
    
    if (lastRow > 1) {
      // Clear all data except headers
      masterSheet.getRange("A2:P" + lastRow).clearContent();
      
      SpreadsheetApp.getUi().alert(
        'Reset Complete',
        'Invoice counter has been reset.\n' +
        'All invoice records have been cleared from the Master sheet.\n' +
        'Next invoice will start from INV-' + new Date().getFullYear() + '-001',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        'No Data',
        'Master sheet is already empty.\n' +
        'No records to clear.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error resetting invoice counter: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function GetInvoiceStats() {
  // Get statistics about invoices in the Master sheet
  try {
    const masterSheet = GetOrCreateWorksheet("Master");
    
    if (!masterSheet) {
      return {
        totalInvoices: 0,
        currentYearInvoices: 0,
        totalValue: 0,
        currentYearValue: 0
      };
    }

    const lastRow = masterSheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        totalInvoices: 0,
        currentYearInvoices: 0,
        totalValue: 0,
        currentYearValue: 0
      };
    }

    const currentYear = new Date().getFullYear();
    const data = masterSheet.getRange("A2:K" + lastRow).getValues();
    
    let totalInvoices = 0;
    let currentYearInvoices = 0;
    let totalValue = 0;
    let currentYearValue = 0;

    data.forEach(row => {
      const invoiceNum = row[0].toString().trim();
      const invoiceValue = parseFloat(row[10]) || 0; // Column K: Total_Invoice_Value
      
      if (invoiceNum) {
        totalInvoices++;
        totalValue += invoiceValue;
        
        if (invoiceNum.indexOf(`INV-${currentYear}-`) === 0) {
          currentYearInvoices++;
          currentYearValue += invoiceValue;
        }
      }
    });

    return {
      totalInvoices,
      currentYearInvoices,
      totalValue,
      currentYearValue
    };
  } catch (error) {
    console.error('Error getting invoice stats:', error);
    return {
      totalInvoices: 0,
      currentYearInvoices: 0,
      totalValue: 0,
      currentYearValue: 0
    };
  }
}

function ShowInvoiceStats() {
  // Display invoice statistics to the user
  try {
    const stats = GetInvoiceStats();
    
    const message = `INVOICE STATISTICS:\n\n` +
                   `📊 TOTAL INVOICES: ${stats.totalInvoices}\n` +
                   `📅 CURRENT YEAR (${new Date().getFullYear()}): ${stats.currentYearInvoices}\n\n` +
                   `💰 TOTAL VALUE: ₹${stats.totalValue.toLocaleString('en-IN', {minimumFractionDigits: 2})}\n` +
                   `💰 CURRENT YEAR VALUE: ₹${stats.currentYearValue.toLocaleString('en-IN', {minimumFractionDigits: 2})}\n\n` +
                   `🔢 NEXT INVOICE NUMBER: ${GetNextInvoiceNumber()}`;

    SpreadsheetApp.getUi().alert('Invoice Statistics', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error showing invoice stats: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
