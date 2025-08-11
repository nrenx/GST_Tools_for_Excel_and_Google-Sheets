/**
 * ===============================================================================
 * MODULE: Module_Warehouse
 * DESCRIPTION: Handles all data management related to the 'warehouse' sheet,
 *              including customer data, HSN codes, and dropdown list setup.
 * ===============================================================================
 */

// ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
// 📋 WORKSHEET CREATION & DATA VALIDATION
// ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

function CreateWarehouseSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Delete existing warehouse sheet if it exists
    const existingSheet = spreadsheet.getSheetByName("warehouse");
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    // Create new warehouse sheet
    const sheet = spreadsheet.insertSheet("warehouse");

    // ===== SECTION 1: HSN/SAC DATA (Columns A-E) =====
    // HSN headers
    const hsnHeaders = ["HSN_Code", "Description", "CGST_Rate", "SGST_Rate", "IGST_Rate"];
    for (let i = 0; i < hsnHeaders.length; i++) {
      sheet.getRange(1, i + 1).setValue(hsnHeaders[i]);
    }

    // Format HSN headers
    sheet.getRange("A1:E1").setFontWeight("bold")
         .setBackground("#2F5061")
         .setFontColor("#FFFFFF")
         .setHorizontalAlignment("center");

    // Add sample HSN data
    const hsnData = [
      ["4401", "Fuel wood, firewood, sawdust, wood waste and scrap", 2.5, 2.5, 5],
      ["4402", "Wood charcoal", 2.5, 2.5, 5],
      ["4403", "Wood in the rough (logs, unprocessed timber)", 9, 9, 18],
      ["4404", "Split poles, pickets, sticks, hoopwood, etc.", 6, 6, 12],
      ["4405", "Wood flour and wood wool", 6, 6, 12],
      ["4406", "Wooden railway or tramway sleepers", 6, 6, 12],
      ["4407", "Wood sawn or chipped", 9, 9, 18],
      ["4408", "Veneered wood and wood continuously shaped", 9, 9, 18],
      ["4409", "Moulded wood, flooring strips", 9, 9, 18],
      ["4410", "Particle board, oriented strand board (OSB), similar boards", 9, 9, 18],
      ["4412", "Plywood, veneered panels, laminated wood", 9, 9, 18],
      ["4413", "Densified wood", 9, 9, 18],
      ["4414", "Wooden frames for mirrors, photos, paintings", 9, 9, 18],
      ["4416", "Wooden barrels, casks, and other cooper's products", 6, 6, 12],
      ["4417", "Wooden tools, tool handles, broom handles", 6, 6, 12],
      ["4418", "Builders' joinery and carpentry of wood (doors, windows, etc.)", 9, 9, 18]
    ];

    // Insert HSN data starting from row 2
    for (let i = 0; i < hsnData.length; i++) {
      for (let j = 0; j < hsnData[i].length; j++) {
        sheet.getRange(i + 2, j + 1).setValue(hsnData[i][j]);
      }
    }

    // ===== SECTION 2: VALIDATION LISTS =====
    // UOM List (Column G)
    sheet.getRange("G1").setValue("UOM_List")
         .setFontWeight("bold")
         .setBackground("#2F5061")
         .setFontColor("#FFFFFF");

    const uomList = ["NOS", "KG", "MT", "CBM", "SQM", "LTR", "PCS", "BOX", "SET", "PAIR"];
    for (let i = 0; i < uomList.length; i++) {
      sheet.getRange(i + 2, 7).setValue(uomList[i]);
    }

    // Transport Mode List (Column H)
    sheet.getRange("H1").setValue("Transport_Mode_List")
         .setFontWeight("bold")
         .setBackground("#2F5061")
         .setFontColor("#FFFFFF");

    const transportList = ["By Lorry", "By Train", "By Air", "By Ship", "By Hand", "Courier", "Self Transport"];
    for (let i = 0; i < transportList.length; i++) {
      sheet.getRange(i + 2, 8).setValue(transportList[i]);
    }

    // State List (Column J) - exactly matching Excel VBA with proper state-code mapping
    sheet.getRange("J1").setValue("State_List")
         .setFontWeight("bold")
         .setBackground("#2F5061")
         .setFontColor("#FFFFFF");

    // State Code List (Column K) - exactly matching Excel VBA
    sheet.getRange("K1").setValue("State_Code_List")
         .setFontWeight("bold")
         .setBackground("#2F5061")
         .setFontColor("#FFFFFF");

    // State and State Code data - exactly matching Excel VBA mapping
    const stateData = [
      ["Jammu and Kashmir", "01"],
      ["Himachal Pradesh", "02"],
      ["Punjab", "03"],
      ["Chandigarh", "04"],
      ["Uttarakhand", "05"],
      ["Haryana", "06"],
      ["Delhi", "07"],
      ["Rajasthan", "08"],
      ["Uttar Pradesh", "09"],
      ["Bihar", "10"],
      ["Sikkim", "11"],
      ["Arunachal Pradesh", "12"],
      ["Nagaland", "13"],
      ["Manipur", "14"],
      ["Mizoram", "15"],
      ["Tripura", "16"],
      ["Meghalaya", "17"],
      ["Assam", "18"],
      ["West Bengal", "19"],
      ["Jharkhand", "20"],
      ["Odisha", "21"],
      ["Chhattisgarh", "22"],
      ["Madhya Pradesh", "23"],
      ["Gujarat", "24"],
      ["Dadra and Nagar Haveli and Daman and Diu", "26"],
      ["Maharashtra", "27"],
      ["Karnataka", "29"],
      ["Goa", "30"],
      ["Lakshadweep", "31"],
      ["Kerala", "32"],
      ["Tamil Nadu", "33"],
      ["Puducherry", "34"],
      ["Andaman and Nicobar Islands", "35"],
      ["Telangana", "36"],
      ["Andhra Pradesh", "37"],
      ["Ladakh", "38"]
    ];

    // Populate state and state code data
    for (let i = 0; i < stateData.length; i++) {
      sheet.getRange(i + 2, 10).setValue(stateData[i][0]); // Column J: State
      sheet.getRange(i + 2, 11).setValue(stateData[i][1]); // Column K: State Code
    }

    // ===== SECTION 3: CUSTOMER MASTER DATA (Columns M-T) =====
    // Customer headers
    const customerHeaders = [
      "Customer_Name", "Address_Line1", "State", "State_Code", "GSTIN", "Phone", "Email", "Contact_Person"
    ];
    
    for (let i = 0; i < customerHeaders.length; i++) {
      sheet.getRange(1, 13 + i).setValue(customerHeaders[i]); // Starting from column M (13)
    }

    // GST Type List (Column X)
    sheet.getRange("X1").setValue("GST_Type")
         .setFontWeight("bold")
         .setBackground("#2F5061")
         .setFontColor("#FFFFFF");
    sheet.getRange("X2").setValue("UNREGISTERED");
    
    // Description List (Column Z)
    sheet.getRange("Z1").setValue("Description")
         .setFontWeight("bold")
         .setBackground("#2F5061")
         .setFontColor("#FFFFFF");
    sheet.getRange("Z2").setValue("Casurina Wood");

    // Set column widths for customer data
    for (let col = 13; col <= 20; col++) { // Columns M to T
      sheet.setColumnWidth(col, 200);
    }

    // Format customer headers
    sheet.getRange("M1:T1").setFontWeight("bold")
         .setBackground("#2F5061")
         .setFontColor("#FFFFFF")
         .setHorizontalAlignment("center")
         .setWrap(true);

    // Add sample customer data
    const sampleCustomers = [
      ["Sample Customer 1", "123 Main Street, City", "Andhra Pradesh", "37", "37XXXXX1234X1Z5", "9876543210", "customer1@email.com", "Contact Person 1"],
      ["Sample Customer 2", "456 Business Ave, Town", "Tamil Nadu", "33", "33XXXXX5678X1Z9", "9876543211", "customer2@email.com", "Contact Person 2"]
    ];

    for (let i = 0; i < sampleCustomers.length; i++) {
      for (let j = 0; j < sampleCustomers[i].length; j++) {
        sheet.getRange(i + 2, 13 + j).setValue(sampleCustomers[i][j]);
      }
    }

    // Set row heights
    sheet.setRowHeight(1, 30);
    
    // Auto-resize columns for better visibility
    sheet.autoResizeColumns(1, 26); // Resize columns A to Z

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error creating warehouse sheet: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function SetupDataValidation(invoiceSheet) {
  try {
    const warehouseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("warehouse");
    
    if (!warehouseSheet) {
      console.error('Warehouse sheet not found for data validation setup');
      return;
    }

    // Set up UOM dropdown for item rows (E18:E21)
    const uomRange = invoiceSheet.getRange("E18:E21");
    const uomValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(warehouseSheet.getRange("G2:G11"), true)
      .setAllowInvalid(true)
      .setHelpText("Select UOM or enter custom value")
      .build();
    uomRange.setDataValidation(uomValidation);

    // Set up Transport Mode dropdown (F7)
    const transportRange = invoiceSheet.getRange("F7");
    const transportValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(warehouseSheet.getRange("H2:H8"), true)
      .setAllowInvalid(true)
      .setHelpText("Select transport mode or enter custom value")
      .build();
    transportRange.setDataValidation(transportValidation);

    // Set up State dropdown for receiver (C15)
    const receiverStateRange = invoiceSheet.getRange("C15");
    const stateValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(warehouseSheet.getRange("J2:J37"), true)
      .setAllowInvalid(true)
      .setHelpText("Select state or enter custom value")
      .build();
    receiverStateRange.setDataValidation(stateValidation);

    // Set up State dropdown for consignee (I15)
    const consigneeStateRange = invoiceSheet.getRange("I15");
    consigneeStateRange.setDataValidation(stateValidation);

    // Set up State Code dropdown for receiver (C16)
    const receiverStateCodeRange = invoiceSheet.getRange("C16");
    const stateCodeValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(warehouseSheet.getRange("K2:K37"), true)
      .setAllowInvalid(true)
      .setHelpText("Select state code or enter custom value")
      .build();
    receiverStateCodeRange.setDataValidation(stateCodeValidation);

    // Set up State Code dropdown for consignee (I16)
    const consigneeStateCodeRange = invoiceSheet.getRange("I16");
    consigneeStateCodeRange.setDataValidation(stateCodeValidation);

  } catch (error) {
    console.error('Error setting up data validation:', error);
  }
}

function SetupCustomerDropdown(invoiceSheet) {
  try {
    const warehouseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("warehouse");
    
    if (!warehouseSheet) {
      console.error('Warehouse sheet not found for customer dropdown setup');
      return;
    }

    // Find the last row with customer data
    const lastRow = warehouseSheet.getLastRow();
    
    if (lastRow > 1) {
      // Set up Customer dropdown for receiver (C12)
      const receiverCustomerRange = invoiceSheet.getRange("C12");
      const customerValidation = SpreadsheetApp.newDataValidation()
        .requireValueInRange(warehouseSheet.getRange(`M2:M${lastRow}`), true)
        .setAllowInvalid(true)
        .setHelpText("Select customer or enter new customer name")
        .build();
      receiverCustomerRange.setDataValidation(customerValidation);

      // Set up Customer dropdown for consignee (I12)
      const consigneeCustomerRange = invoiceSheet.getRange("I12");
      consigneeCustomerRange.setDataValidation(customerValidation);
    }

  } catch (error) {
    console.error('Error setting up customer dropdown:', error);
  }
}

function SetupHSNDropdown(invoiceSheet) {
  try {
    const warehouseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("warehouse");
    
    if (!warehouseSheet) {
      console.error('Warehouse sheet not found for HSN dropdown setup');
      return;
    }

    // Set up HSN Code dropdown for item rows (C18:C21)
    const hsnRange = invoiceSheet.getRange("C18:C21");
    const hsnValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(warehouseSheet.getRange("A2:A17"), true)
      .setAllowInvalid(true)
      .setHelpText("Select HSN code or enter custom value")
      .build();
    hsnRange.setDataValidation(hsnValidation);

  } catch (error) {
    console.error('Error setting up HSN dropdown:', error);
  }
}
