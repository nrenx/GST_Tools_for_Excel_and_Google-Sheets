/**
 * ===============================================================================
 * MODULE: Module_InvoiceEvents
 * DESCRIPTION: Handles all button clicks, event handlers, and user interactions
 *              on the invoice worksheet.
 * ===============================================================================
 */

// ████████████████████████████████████████████████████████████████████████████████
// 🔘 BUTTON FUNCTIONS - DAILY OPERATIONS
// ████████████████████████████████████████████████████████████████████████████████
// These functions are designed to be assigned to buttons for daily use.

function AddCustomerToWarehouseButton() {
  // Button function: Capture customer details from current invoice and save to warehouse
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const invoiceSheet = spreadsheet.getSheetByName("GST_Tax_Invoice_for_interstate");
    const warehouseSheet = spreadsheet.getSheetByName("warehouse");

    if (!invoiceSheet || !warehouseSheet) {
      SpreadsheetApp.getUi().alert('Error', 'Required worksheets not found. Please run setup first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Get customer details from invoice
    const customerName = invoiceSheet.getRange("C12").getValue().toString().trim();
    const address = (invoiceSheet.getRange("C13").getValue().toString() + " " + 
                    invoiceSheet.getRange("C14").getValue().toString() + " " + 
                    invoiceSheet.getRange("C15").getValue().toString()).trim();
    const gstin = invoiceSheet.getRange("C16").getValue().toString().trim();
    const stateCode = invoiceSheet.getRange("C10").getValue().toString().trim();

    // Validate required fields
    if (!customerName) {
      SpreadsheetApp.getUi().alert('Missing Information', 'Please enter customer name before adding to warehouse.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Check for duplicates in warehouse (Customer section - columns M-T)
    const lastRow = warehouseSheet.getLastRow();
    const customerData = warehouseSheet.getRange("M2:M" + lastRow).getValues();

    for (let i = 0; i < customerData.length; i++) {
      if (customerData[i][0].toString().toUpperCase().trim() === customerName.toUpperCase()) {
        SpreadsheetApp.getUi().alert('Duplicate Customer', `Customer '${customerName}' already exists in warehouse.`, SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
    }

    // Add new customer to next available row
    const newRow = lastRow + 1;
    warehouseSheet.getRange(`M${newRow}`).setValue(customerName);     // Column M: Customer Name
    warehouseSheet.getRange(`N${newRow}`).setValue(address);          // Column N: Address
    warehouseSheet.getRange(`O${newRow}`).setValue("");               // Column O: State (empty for now)
    warehouseSheet.getRange(`P${newRow}`).setValue(stateCode);        // Column P: State Code
    warehouseSheet.getRange(`Q${newRow}`).setValue(gstin);            // Column Q: GSTIN
    warehouseSheet.getRange(`R${newRow}`).setValue("");               // Column R: Phone (empty)
    warehouseSheet.getRange(`S${newRow}`).setValue("");               // Column S: Email (empty)
    warehouseSheet.getRange(`T${newRow}`).setValue("");               // Column T: Contact Person (empty)

    SpreadsheetApp.getUi().alert('Customer Added', `Customer '${customerName}' added successfully to warehouse!`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error adding customer: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function AddNewItemRowButton() {
  // Button function: Add new item row after existing item rows with clean layout
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GST_Tax_Invoice_for_interstate");
  AddNewItemRow(sheet);
}

function NewInvoiceButton() {
  // Button function: Generate a fresh invoice with next sequential number and cleared fields
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GST_Tax_Invoice_for_interstate");
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Error', 'Invoice sheet not found. Please run setup first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Confirm creating new invoice
    const response = SpreadsheetApp.getUi().alert(
      'Confirm New Invoice',
      'Create a new invoice?\nAll current data will be cleared and a new invoice number will be generated.',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (response !== SpreadsheetApp.getUi().Button.YES) return;

    // Generate next sequential invoice number
    const nextInvoiceNumber = GetNextInvoiceNumber();

    // Clear and set invoice number (C7) with new sequential number
    sheet.getRange("C7").setValue(nextInvoiceNumber)
         .setFontWeight("bold").setFontColor("#DC143C")
         .setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Set current date for Invoice Date (C8) and Date of Supply (F9, G9)
    const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    sheet.getRange("C8").setValue(currentDate)
         .setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");
    
    sheet.getRange("F9").setValue(currentDate)
         .setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");
    
    sheet.getRange("G9").setValue(currentDate)
         .setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");

    // Reset state code to default (C10)
    sheet.getRange("C10").setValue("37")
         .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Clear all customer details (handle merged cells properly)
    sheet.getRange("C12:F15").clearContent(); // Clear Receiver details, preserving formulas in row 16
    sheet.getRange("I12:K15").clearContent(); // Clear Consignee details, preserving formulas in row 16
    sheet.getRange("F7").setValue("By Lorry");   // Reset Transport Mode
    sheet.getRange("F8").setValue("");           // Clear Vehicle Number
    sheet.getRange("F10").setValue("");          // Clear Place of Supply

    // Clear item table data (rows 18-21, keep headers and formulas)
    sheet.getRange("A18:F21").clearContent();
    // Reset first Sr.No.
    sheet.getRange("A18").setValue(1);

    // Clear tax summary section (handle merged cells properly)
    sheet.getRange("K23").clearContent();  // Total Before Tax
    sheet.getRange("K24").clearContent();  // CGST
    sheet.getRange("K25").clearContent();  // SGST
    sheet.getRange("K26").clearContent();  // IGST
    sheet.getRange("K27").clearContent();  // Round Off
    sheet.getRange("K28").clearContent();  // Total Tax Amount
    sheet.getRange("K29").clearContent();  // Total Amount

    // Clear amount in words
    sheet.getRange("A30").clearContent();

    // Clear other optional fields
    sheet.getRange("J7").setValue("");   // Challan No.
    sheet.getRange("J9").setValue("");   // L.R Number
    sheet.getRange("J10").setValue("");  // P.O Number

    SpreadsheetApp.getUi().alert('New Invoice Created', `New invoice created with number: ${nextInvoiceNumber}`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error creating new invoice: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function SaveInvoiceButton() {
  // Button function: Save complete invoice record to Master sheet for auditing
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const invoiceSheet = spreadsheet.getSheetByName("GST_Tax_Invoice_for_interstate");
    const masterSheet = spreadsheet.getSheetByName("Master");

    if (!invoiceSheet || !masterSheet) {
      SpreadsheetApp.getUi().alert('Error', 'Required worksheets not found. Please run setup first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Get invoice data
    const invoiceNumber = invoiceSheet.getRange("C7").getValue().toString().trim();
    const invoiceDate = invoiceSheet.getRange("C8").getValue();
    const customerName = invoiceSheet.getRange("C12").getValue().toString().trim();
    const customerGSTIN = invoiceSheet.getRange("C14").getValue().toString().trim();
    const customerState = invoiceSheet.getRange("C15").getValue().toString().trim();
    const customerStateCode = invoiceSheet.getRange("C16").getValue().toString().trim();

    // Validate required fields
    if (!invoiceNumber || !customerName) {
      SpreadsheetApp.getUi().alert('Missing Information', 'Please ensure invoice number and customer name are filled.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Get tax amounts
    const totalTaxableValue = invoiceSheet.getRange("K22").getValue() || 0;
    const igstRate = invoiceSheet.getRange("I18").getValue() || 0; // Get from first item row
    const igstAmount = invoiceSheet.getRange("K25").getValue() || 0;
    const totalTaxAmount = invoiceSheet.getRange("K27").getValue() || 0;
    const totalInvoiceValue = invoiceSheet.getRange("K28").getValue() || 0;

    // Get item details (combine all items)
    let hsnCodes = [];
    let itemDescriptions = [];
    let quantities = [];
    let uoms = [];

    for (let row = 18; row <= 21; row++) {
      const hsn = invoiceSheet.getRange(`C${row}`).getValue().toString().trim();
      const desc = invoiceSheet.getRange(`B${row}`).getValue().toString().trim();
      const qty = invoiceSheet.getRange(`D${row}`).getValue();
      const uom = invoiceSheet.getRange(`E${row}`).getValue().toString().trim();

      if (hsn || desc) {
        hsnCodes.push(hsn);
        itemDescriptions.push(desc);
        quantities.push(qty);
        uoms.push(uom);
      }
    }

    // Check for duplicate invoice number
    const lastRow = masterSheet.getLastRow();
    if (lastRow > 1) {
      const existingInvoices = masterSheet.getRange("A2:A" + lastRow).getValues();
      for (let i = 0; i < existingInvoices.length; i++) {
        if (existingInvoices[i][0].toString().trim() === invoiceNumber) {
          const response = SpreadsheetApp.getUi().alert(
            'Duplicate Invoice',
            `Invoice ${invoiceNumber} already exists. Do you want to update it?`,
            SpreadsheetApp.getUi().ButtonSet.YES_NO
          );
          if (response !== SpreadsheetApp.getUi().Button.YES) return;
          
          // Update existing record
          const updateRow = i + 2; // +2 because array is 0-based and we start from row 2
          UpdateMasterRecord(masterSheet, updateRow, invoiceNumber, invoiceDate, customerName, customerGSTIN, 
                           customerState, customerStateCode, totalTaxableValue, igstRate, igstAmount, 
                           totalTaxAmount, totalInvoiceValue, hsnCodes, itemDescriptions, quantities, uoms);
          return;
        }
      }
    }

    // Add new record
    const newRow = lastRow + 1;
    UpdateMasterRecord(masterSheet, newRow, invoiceNumber, invoiceDate, customerName, customerGSTIN, 
                     customerState, customerStateCode, totalTaxableValue, igstRate, igstAmount, 
                     totalTaxAmount, totalInvoiceValue, hsnCodes, itemDescriptions, quantities, uoms);

    SpreadsheetApp.getUi().alert('Invoice Saved', `Invoice ${invoiceNumber} saved successfully to Master sheet!`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error saving invoice: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function UpdateMasterRecord(masterSheet, row, invoiceNumber, invoiceDate, customerName, customerGSTIN, 
                          customerState, customerStateCode, totalTaxableValue, igstRate, igstAmount, 
                          totalTaxAmount, totalInvoiceValue, hsnCodes, itemDescriptions, quantities, uoms) {
  try {
    masterSheet.getRange(`A${row}`).setValue(invoiceNumber);
    masterSheet.getRange(`B${row}`).setValue(invoiceDate);
    masterSheet.getRange(`C${row}`).setValue(customerName);
    masterSheet.getRange(`D${row}`).setValue(customerGSTIN);
    masterSheet.getRange(`E${row}`).setValue(customerState);
    masterSheet.getRange(`F${row}`).setValue(customerStateCode);
    masterSheet.getRange(`G${row}`).setValue(totalTaxableValue);
    masterSheet.getRange(`H${row}`).setValue(igstRate);
    masterSheet.getRange(`I${row}`).setValue(igstAmount);
    masterSheet.getRange(`J${row}`).setValue(totalTaxAmount);
    masterSheet.getRange(`K${row}`).setValue(totalInvoiceValue);
    masterSheet.getRange(`L${row}`).setValue(hsnCodes.join(", "));
    masterSheet.getRange(`M${row}`).setValue(itemDescriptions.join(", "));
    masterSheet.getRange(`N${row}`).setValue(quantities.join(", "));
    masterSheet.getRange(`O${row}`).setValue(uoms.join(", "));
    masterSheet.getRange(`P${row}`).setValue(new Date());
  } catch (error) {
    console.error('Error updating master record:', error);
  }
}

function PrintAsPDFButton() {
  // Button function: Export ONLY the invoice sheet as PDF to Google Drive (exactly matching Excel VBA)
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const invoiceSheet = spreadsheet.getSheetByName("GST_Tax_Invoice_for_interstate");

    if (!invoiceSheet) {
      SpreadsheetApp.getUi().alert('Error', 'Invoice sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Get invoice number for filename
    const invoiceNumber = invoiceSheet.getRange("C7").getValue().toString().trim();
    const customerName = invoiceSheet.getRange("C12").getValue().toString().trim();

    if (!invoiceNumber) {
      SpreadsheetApp.getUi().alert('Missing Information', 'Please ensure invoice number is filled.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Create PDF filename - exactly matching Excel VBA format
    const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const cleanCustomerName = customerName.replace(/[^a-zA-Z0-9]/g, "_").substring(0, 20);
    const filename = `${invoiceNumber}_${cleanCustomerName}_${currentDate}.pdf`;

    // Get the target Google Drive folder
    const folderId = "1boyjaNQVZMZ6Gk_bRsTY7B0D7Lre_1r7"; // Your specified folder ID
    let targetFolder;

    try {
      targetFolder = DriveApp.getFolderById(folderId);
    } catch (e) {
      SpreadsheetApp.getUi().alert('Drive Error', 'Cannot access the specified Google Drive folder. Please check permissions.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // CRITICAL FIX: Hide all sheets except the invoice sheet before PDF export
    const allSheets = spreadsheet.getSheets();
    const sheetsToHide = [];

    allSheets.forEach(sheet => {
      if (sheet.getName() !== "GST_Tax_Invoice_for_interstate") {
        if (!sheet.isSheetHidden()) {
          sheet.hideSheet();
          sheetsToHide.push(sheet);
        }
      }
    });

    // Create PDF blob with only the invoice sheet visible
    const pdfBlob = spreadsheet.getAs('application/pdf');
    pdfBlob.setName(filename);

    // Restore visibility of hidden sheets
    sheetsToHide.forEach(sheet => {
      sheet.showSheet();
    });

    // Save to Google Drive
    const pdfFile = targetFolder.createFile(pdfBlob);

    SpreadsheetApp.getUi().alert(
      'PDF Created',
      `PDF exported successfully!\nFilename: ${filename}\nSaved to: GST Invoices folder\nFile ID: ${pdfFile.getId()}\n\nNote: PDF contains ONLY the invoice sheet.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error creating PDF: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function PrintButton() {
  // Button function: Save PDF and provide print instructions
  try {
    // First save as PDF
    PrintAsPDFButton();

    // Provide print instructions
    SpreadsheetApp.getUi().alert(
      'Print Instructions',
      'PDF has been saved to Google Drive.\n\n' +
      'To print:\n' +
      '1. Open the PDF from Google Drive\n' +
      '2. Click the print icon or press Ctrl+P\n' +
      '3. Select your printer and print settings\n' +
      '4. Click Print',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error in print function: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ████████████████████████████████████████████████████████████████████████████████
// 🔧 BUTTON CREATION AND MANAGEMENT
// ████████████████████████████████████████████████████████████████████████████████

function CreateInvoiceButtons(sheet) {
  try {
    // Note: Google Sheets doesn't support VBA-style buttons
    // Instead, we'll create a menu system or use the script editor
    // Users can run functions directly from the script editor or create custom menus

    // Create custom menu for easy access to functions
    CreateCustomMenu();

    // Add instructions in a comment or separate cell
    sheet.getRange("P1").setValue("Use Extensions > Apps Script menu for invoice functions")
         .setFontSize(8).setFontColor("#666666").setWrap(true);

  } catch (error) {
    console.error('Error creating invoice buttons:', error);
  }
}

function CreateCustomMenu() {
  try {
    const ui = SpreadsheetApp.getUi();

    ui.createMenu('GST Invoice System')
      .addSubMenu(ui.createMenu('Setup')
        .addItem('Quick Setup', 'QuickSetup')
        .addItem('Full Setup', 'StartGSTSystem')
        .addItem('Show Available Functions', 'ShowAvailableFunctions'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Invoice Operations')
        .addItem('New Invoice', 'NewInvoiceButton')
        .addItem('Add New Item Row', 'AddNewItemRowButton')
        .addItem('Save Invoice', 'SaveInvoiceButton')
        .addItem('Export as PDF', 'PrintAsPDFButton')
        .addItem('Print Invoice', 'PrintButton'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Customer Management')
        .addItem('Add Customer to Warehouse', 'AddCustomerToWarehouseButton'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Utilities')
        .addItem('Verify Validation Settings', 'VerifyValidationSettings')
        .addItem('Auto-fill Consignee from Receiver', 'AutoFillConsigneeFromReceiverButton'))
      .addToUi();
  } catch (error) {
    console.error('Error creating custom menu:', error);
  }
}

function AutoFillConsigneeFromReceiverButton() {
  // Button function: Auto-fill consignee details from receiver details
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GST_Tax_Invoice_for_interstate");
    if (sheet) {
      AutoFillConsigneeFromReceiver(sheet);
      SpreadsheetApp.getUi().alert('Success', 'Consignee details filled from receiver information.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error auto-filling consignee: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ████████████████████████████████████████████████████████████████████████████████
// 🎯 EVENT HANDLERS AND TRIGGERS
// ████████████████████████████████████████████████████████████████████████████████

function onOpen() {
  // This function runs automatically when the spreadsheet is opened
  CreateCustomMenu();
}

function onEdit(e) {
  // This function runs automatically when any cell is edited - exactly matching Excel VBA functionality
  try {
    const sheet = e.source.getActiveSheet();
    const range = e.range;

    // Only process edits on the invoice sheet
    if (sheet.getName() !== "GST_Tax_Invoice_for_interstate") return;

    const row = range.getRow();
    const col = range.getColumn();
    const editedValue = range.getValue();

    // CRITICAL FIX: State code auto-fill functionality (exactly matching Excel VBA)
    // When state is selected in receiver section (C15), auto-fill state code (C16)
    if (row === 15 && col === 3 && editedValue) {
      const warehouseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("warehouse");
      if (warehouseSheet) {
        // Use VLOOKUP to find the corresponding state code
        try {
          const stateCode = warehouseSheet.getRange("J2:K37").getValues()
            .find(stateRow => stateRow[0] === editedValue);
          if (stateCode) {
            sheet.getRange("C16").setValue(stateCode[1]);
          }
        } catch (error) {
          console.error('Error in state code lookup:', error);
        }
      }
    }

    // When state is selected in consignee section (I15), auto-fill state code (I16)
    if (row === 15 && col === 9 && editedValue) {
      const warehouseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("warehouse");
      if (warehouseSheet) {
        try {
          const stateCode = warehouseSheet.getRange("J2:K37").getValues()
            .find(stateRow => stateRow[0] === editedValue);
          if (stateCode) {
            sheet.getRange("I16").setValue(stateCode[1]);
          }
        } catch (error) {
          console.error('Error in consignee state code lookup:', error);
        }
      }
    }

    // HSN code auto-fill functionality (exactly matching Excel VBA)
    // When HSN code is selected, auto-fill IGST rate
    if (row >= 18 && row <= 21 && col === 3 && editedValue) {
      const warehouseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("warehouse");
      if (warehouseSheet) {
        try {
          const hsnData = warehouseSheet.getRange("A2:E17").getValues()
            .find(hsnRow => hsnRow[0] === editedValue);
          if (hsnData) {
            sheet.getRange(row, 9).setValue(hsnData[4]); // Column I: IGST Rate
          }
        } catch (error) {
          console.error('Error in HSN code lookup:', error);
        }
      }
    }

    // Auto-calculate formulas when item data is edited (rows 18-21)
    if (row >= 18 && row <= 21 && (col === 4 || col === 6)) { // Quantity or Rate columns
      // Trigger recalculation of amount formulas
      const amountCell = sheet.getRange(row, 7); // Column G: Amount
      const taxableValueCell = sheet.getRange(row, 8); // Column H: Taxable Value
      const igstAmountCell = sheet.getRange(row, 10); // Column J: IGST Amount
      const totalAmountCell = sheet.getRange(row, 11); // Column K: Total Amount

      // Force recalculation by setting formulas again
      amountCell.setFormula(`=IF(AND(D${row}<>"",F${row}<>""),D${row}*F${row},"")`);
      taxableValueCell.setFormula(`=G${row}`);
      igstAmountCell.setFormula(`=IF(AND(H${row}<>"",I${row}<>""),H${row}*I${row}/100,"")`);
      totalAmountCell.setFormula(`=IF(H${row}<>"",H${row}+J${row},"")`);

      // Update summary calculations
      UpdateMultiItemTaxCalculations(sheet);
    }

  } catch (error) {
    console.error('Error in onEdit trigger:', error);
  }
}
