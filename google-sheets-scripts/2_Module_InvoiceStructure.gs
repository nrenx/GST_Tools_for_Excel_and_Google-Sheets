/**
 * ===============================================================================
 * MODULE: Module_InvoiceStructure
 * DESCRIPTION: Handles the creation, formatting, and layout of the invoice sheet,
 *              as well as its core formulas and data population logic.
 * ===============================================================================
 */

// ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
// 📋 WORKSHEET CREATION FUNCTIONS
// ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

function CreateInvoiceSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Try to get the sheet
    let sheet = spreadsheet.getSheetByName("GST_Tax_Invoice_for_interstate");
    
    // If the sheet doesn't exist, create it. If it exists, clear it completely.
    if (!sheet) {
      sheet = spreadsheet.insertSheet("GST_Tax_Invoice_for_interstate");
    } else {
      // Complete cleanup of existing sheet
      sheet.clear();
      sheet.clearFormats();
    }

    // Set column widths
    sheet.setColumnWidth(1, 60);   // Column A - Sr.No.
    sheet.setColumnWidth(2, 144);  // Column B - Description of Goods/Services
    sheet.setColumnWidth(3, 144);  // Column C - HSN/SAC Code
    sheet.setColumnWidth(4, 108);  // Column D - Quantity
    sheet.setColumnWidth(5, 84);   // Column E - UOM
    sheet.setColumnWidth(6, 120);  // Column F - Rate
    sheet.setColumnWidth(7, 168);  // Column G - Amount
    sheet.setColumnWidth(8, 120);  // Column H - Taxable Value
    sheet.setColumnWidth(9, 72);   // Column I - IGST Rate
    sheet.setColumnWidth(10, 120); // Column J - IGST Amount
    sheet.setColumnWidth(11, 144); // Column K - Total Amount

    // Set default font for all cells
    const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    range.setFontFamily("Segoe UI")
         .setFontSize(11)
         .setFontColor("#1a1a1a");

    // Create header sections with premium professional styling
    CreateHeaderRow(sheet, 1, "A1:K1", "ORIGINAL", 12, true, "#2F5061", "#FFFFFF", 25);
    CreateHeaderRow(sheet, 2, "A2:K2", "KAVERI TRADERS", 24, true, "#2F5061", "#FFFFFF", 37);
    CreateHeaderRow(sheet, 3, "A3:K3", "191, Guduru, Pagadalapalli, Idulapalli, Tirupati, Andhra Pradesh - 524409", 11, true, "#F5F5F5", "#1a1a1a", 27);
    CreateHeaderRow(sheet, 4, "A4:K4", "GSTIN: 37HERPB7733F1Z5", 14, true, "#F5F5F5", "#1a1a1a", 27);
    CreateHeaderRow(sheet, 5, "A5:K5", "Email: kotidarisetty7777@gmail.com", 11, true, "#F5F5F5", "#1a1a1a", 25);

    // Row 6: TAX-INVOICE header
    CreateHeaderRow(sheet, 6, "A6:G6", "TAX-INVOICE", 22, true, "#F0F0F0", "#000000", 28);
    CreateHeaderRow(sheet, 6, "H6:K6", "Original for Recipient\nDuplicate for Supplier/Transporter\nTriplicate for Supplier", 9, true, "#FAFAFA", "#000000", 45);

    // Enable text wrapping for the right section and ensure center alignment for TAX-INVOICE
    sheet.getRange("A6:G6").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange("H6:K6").setWrap(true);

    // Create invoice details section
    CreateInvoiceDetailsSection(sheet);
    
    // Create party details section
    CreatePartyDetailsSection(sheet);
    
    // Create item table section
    CreateItemTableSection(sheet);
    
    // Create tax summary section
    CreateTaxSummarySection(sheet);
    
    // Create bottom section
    CreateBottomSection(sheet);
    
    // Create buttons
    CreateInvoiceButtons(sheet);
    
    // Auto-populate initial fields
    AutoPopulateInvoiceFields(sheet);

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error creating invoice sheet: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function CreateHeaderRow(sheet, row, range, text, fontSize, bold, bgColor, fontColor, height) {
  try {
    const headerRange = sheet.getRange(range);
    headerRange.merge()
             .setValue(text)
             .setFontSize(fontSize)
             .setFontWeight(bold ? "bold" : "normal")
             .setBackground(bgColor)
             .setFontColor(fontColor)
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle")
             .setBorder(true, true, true, true, false, false, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
    
    sheet.setRowHeight(row, height);
  } catch (error) {
    console.error('Error creating header row:', error);
  }
}

function CreateInvoiceDetailsSection(sheet) {
  try {
    // Row 7: Invoice No., Transport Mode, Challan No.
    sheet.getRange("A7:B7").merge().setValue("Invoice No.")
         .setFontWeight("bold").setHorizontalAlignment("left")
         .setBackground("#F5F5F5").setFontColor("#1a1a1a");
    
    sheet.getRange("C7").setValue("")
         .setFontWeight("bold").setFontColor("#DC143C")
         .setHorizontalAlignment("center").setVerticalAlignment("middle");

    sheet.getRange("D7:E7").merge().setValue("Transport Mode")
         .setFontWeight("bold").setHorizontalAlignment("left")
         .setBackground("#F5F5F5").setFontColor("#1a1a1a");
    
    sheet.getRange("F7:G7").merge().setValue("By Lorry")
         .setHorizontalAlignment("left");

    sheet.getRange("H7:I7").merge().setValue("Challan No.")
         .setFontWeight("bold").setHorizontalAlignment("left")
         .setBackground("#F5F5F5").setFontColor("#1a1a1a");
    
    sheet.getRange("J7:K7").merge().setValue("")
         .setHorizontalAlignment("left");

    // Row 8: Invoice Date, Vehicle Number, Transporter Name
    sheet.getRange("A8:B8").merge().setValue("Invoice Date")
         .setFontWeight("bold").setHorizontalAlignment("left")
         .setBackground("#F5F5F5").setFontColor("#1a1a1a");
    
    sheet.getRange("C8").setValue("")
         .setFontWeight("bold").setHorizontalAlignment("left");

    sheet.getRange("D8:E8").merge().setValue("Vehicle Number")
         .setFontWeight("bold").setHorizontalAlignment("left")
         .setBackground("#F5F5F5").setFontColor("#1a1a1a");
    
    sheet.getRange("F8:G8").merge().setValue("")
         .setHorizontalAlignment("left");

    sheet.getRange("H8:I8").merge().setValue("Transporter Name")
         .setFontWeight("bold").setHorizontalAlignment("left")
         .setBackground("#F5F5F5").setFontColor("#1a1a1a");
    
    sheet.getRange("J8:K8").merge().setValue("NARENDRA")
         .setHorizontalAlignment("left");

    // Row 9: State, Date of Supply, L.R Number
    sheet.getRange("A9:B9").merge().setValue("State")
         .setFontWeight("bold").setHorizontalAlignment("left")
         .setBackground("#F5F5F5").setFontColor("#1a1a1a");
    
    sheet.getRange("C9").setValue("Andhra Pradesh")
         .setHorizontalAlignment("left").setFontSize(10);

    sheet.getRange("D9:E9").merge().setValue("Date of Supply")
         .setFontWeight("bold").setHorizontalAlignment("left")
         .setBackground("#F5F5F5").setFontColor("#1a1a1a");
    
    sheet.getRange("F9:G9").merge().setValue("")
         .setHorizontalAlignment("left");

    sheet.getRange("H9:I9").merge().setValue("L.R Number")
         .setFontWeight("bold").setHorizontalAlignment("left")
         .setBackground("#F5F5F5").setFontColor("#1a1a1a");
    
    sheet.getRange("J9:K9").merge().setValue("")
         .setHorizontalAlignment("left");

    // Row 10: State Code, Place of Supply, P.O Number
    sheet.getRange("A10:B10").merge().setValue("State Code")
         .setFontWeight("bold").setHorizontalAlignment("left")
         .setBackground("#F5F5F5").setFontColor("#1a1a1a");
    
    sheet.getRange("C10").setValue("37")
         .setHorizontalAlignment("left");

    sheet.getRange("D10:E10").merge().setValue("Place of Supply")
         .setFontWeight("bold").setHorizontalAlignment("left")
         .setBackground("#F5F5F5").setFontColor("#1a1a1a");
    
    sheet.getRange("F10:G10").merge().setValue("")
         .setHorizontalAlignment("left");

    sheet.getRange("H10:I10").merge().setValue("P.O Number")
         .setFontWeight("bold").setHorizontalAlignment("left")
         .setBackground("#F5F5F5").setFontColor("#1a1a1a");
    
    sheet.getRange("J10:K10").merge().setValue("")
         .setHorizontalAlignment("left");

    // Apply borders and formatting with professional color
    sheet.getRange("A7:K10").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
    
    for (let i = 7; i <= 10; i++) {
      sheet.setRowHeight(i, 28);
    }
  } catch (error) {
    console.error('Error creating invoice details section:', error);
  }
}

function CreatePartyDetailsSection(sheet) {
  try {
    // Party Details Headers
    CreateHeaderRow(sheet, 11, "A11:F11", "Details of Receiver (Billed to)", 11, true, "#F5F5F5", "#1a1a1a", 26);
    CreateHeaderRow(sheet, 11, "G11:K11", "Details of Consignee (Shipped to)", 11, true, "#F5F5F5", "#1a1a1a", 26);

    // Set center alignment for row 11 content
    sheet.getRange("A11:F11").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange("G11:K11").setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Create party detail fields (rows 12-16)
    const partyFields = [
      { row: 12, label: "Name:" },
      { row: 13, label: "Address:" },
      { row: 14, label: "GSTIN:" },
      { row: 15, label: "State:" },
      { row: 16, label: "State Code:" }
    ];

    partyFields.forEach(field => {
      // Receiver section
      sheet.getRange(`A${field.row}:B${field.row}`).merge().setValue(field.label)
           .setFontWeight("bold").setHorizontalAlignment("left")
           .setBackground("#F5F5F5").setFontColor("#1a1a1a");
      
      sheet.getRange(`C${field.row}:F${field.row}`).merge().setValue("")
           .setHorizontalAlignment("left");

      // Consignee section
      sheet.getRange(`G${field.row}:H${field.row}`).merge().setValue(field.label)
           .setFontWeight("bold").setHorizontalAlignment("left")
           .setBackground("#F5F5F5").setFontColor("#1a1a1a");

      sheet.getRange(`I${field.row}:K${field.row}`).merge().setValue("")
           .setHorizontalAlignment("left");
    });

    // Add VLOOKUP formulas for state code auto-fill (exactly matching Excel VBA)
    // Row 16: State Code fields with VLOOKUP formulas
    sheet.getRange("C16").setFormula("=VLOOKUP(C15, warehouse!J2:K37, 2, FALSE)")
         .setHorizontalAlignment("left");

    sheet.getRange("I16").setFormula("=VLOOKUP(I15, warehouse!J2:K37, 2, FALSE)")
         .setHorizontalAlignment("left");

    // Apply borders and formatting
    sheet.getRange("A11:K16").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
    
    for (let i = 11; i <= 16; i++) {
      sheet.setRowHeight(i, 26);
    }
  } catch (error) {
    console.error('Error creating party details section:', error);
  }
}

function CreateItemTableSection(sheet) {
  try {
    // Item table headers (row 17) - exactly matching Excel VBA
    const itemHeaders = [
      "Sr.No.", "Description of Goods/Services", "HSN/SAC Code", "Quantity", "UOM",
      "Rate (Rs.)", "Amount (Rs.)", "Taxable Value (Rs.)", "IGST Rate (%)", "IGST Amount (Rs.)", "Total Amount (Rs.)"
    ];

    for (let i = 0; i < itemHeaders.length; i++) {
      sheet.getRange(17, i + 1).setValue(itemHeaders[i])
           .setFontWeight("bold").setFontSize(10)
           .setBackground("#F5F5F5").setFontColor("#1a1a1a")
           .setHorizontalAlignment("center").setVerticalAlignment("middle")
           .setWrap(true);
    }

    // Set row height for header
    sheet.setRowHeight(17, 58);

    // Create first item row (18) with sample data
    const itemData = ["1", "Casuarina Wood", "", "", "", "", "", "", "", "", ""];
    for (let i = 0; i < itemData.length; i++) {
      const cell = sheet.getRange(18, i + 1);
      cell.setValue(itemData[i])
          .setFontSize(10)
          .setBackground("#FAFAFA");

      // Set alignment based on column
      if (i === 0 || i === 2 || i === 3 || i === 4) {
        cell.setHorizontalAlignment("center");
      } else if (i === 1) {
        cell.setHorizontalAlignment("left");
      } else if (i >= 5) {
        cell.setHorizontalAlignment("right").setFontWeight("bold");
      }
    }
    sheet.setRowHeight(18, 35);

    // Create empty rows 19-21 with alternating colors
    for (let row = 19; row <= 21; row++) {
      const bgColor = (row % 2 === 0) ? "#FAFAFA" : "#FFFFFF";
      sheet.getRange(`A${row}:K${row}`).setBackground(bgColor);
      sheet.setRowHeight(row, 30);
    }

    // Apply borders to entire item table
    sheet.getRange("A17:K21").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

    // Setup automatic tax calculation formulas
    SetupTaxCalculationFormulas(sheet);

  } catch (error) {
    console.error('Error creating item table section:', error);
  }
}

function CreateTaxSummarySection(sheet) {
  try {
    // Row 22: Total Quantity Section (exactly matching Excel VBA)
    sheet.getRange("A22:C22").merge().setValue("Total Quantity")
         .setFontWeight("bold").setHorizontalAlignment("center")
         .setVerticalAlignment("bottom").setBackground("#EAEAEA").setFontColor("#1a1a1a");

    sheet.getRange("D22").setValue("")
         .setFontWeight("bold").setHorizontalAlignment("center")
         .setBackground("#EAEAEA");

    sheet.getRange("E22:F22").merge().setValue("Sub Total:")
         .setFontWeight("bold").setHorizontalAlignment("right")
         .setBackground("#EAEAEA").setFontColor("#1a1a1a");

    // Individual cells for amounts
    ["G22", "H22"].forEach(cell => {
      sheet.getRange(cell).setValue("")
           .setFontWeight("bold").setHorizontalAlignment("right")
           .setBackground("#EAEAEA");
    });

    sheet.getRange("I22:J22").merge().setValue("")
         .setFontWeight("bold").setHorizontalAlignment("right");

    sheet.getRange("K22").setValue("")
         .setFontWeight("bold").setHorizontalAlignment("right");

    // Apply formatting to entire row 22
    sheet.getRange("A22:K22").setBackground("#EAEAEA")
         .setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(22, 26);

    // Row 23-25: Total Invoice Amount in Words Section
    sheet.getRange("A23:G23").merge().setValue("Total Invoice Amount in Words")
         .setFontWeight("bold").setFontSize(13)
         .setHorizontalAlignment("center").setBackground("#FFFF00");
    sheet.setRowHeight(23, 25);

    // Rows 24-25: Amount in words content (merged across 2 rows)
    sheet.getRange("A24:G25").merge().setValue("")
         .setFontWeight("bold").setFontSize(15)
         .setHorizontalAlignment("center").setVerticalAlignment("middle")
         .setBackground("#FFFFE6").setWrap(true);
    sheet.setRowHeight(24, 25);
    sheet.setRowHeight(25, 25);

    // Tax summary on the right (columns H-K, rows 23-25)
    // Row 23: Total Amount Before Tax
    sheet.getRange("H23:J23").merge().setValue("Total Amount Before Tax:")
         .setFontWeight("bold").setFontSize(11)
         .setHorizontalAlignment("left").setBackground("#F5F5F5").setFontColor("#1a1a1a");

    sheet.getRange("K23").setValue("")
         .setFontWeight("bold").setHorizontalAlignment("right")
         .setBackground("#D8DEE9");

    // Row 24: CGST
    sheet.getRange("H24:J24").merge().setValue("CGST :")
         .setFontWeight("bold").setFontSize(11)
         .setHorizontalAlignment("left").setBackground("#F5F5F5").setFontColor("#1a1a1a");

    sheet.getRange("K24").setValue("")
         .setFontWeight("bold").setHorizontalAlignment("right")
         .setBackground("#D8DEE9");

    // Row 25: SGST
    sheet.getRange("H25:J25").merge().setValue("SGST :")
         .setFontWeight("bold").setFontSize(11)
         .setHorizontalAlignment("left").setBackground("#F5F5F5").setFontColor("#1a1a1a");

    sheet.getRange("K25").setValue("")
         .setFontWeight("bold").setHorizontalAlignment("right")
         .setBackground("#D8DEE9");

    // Apply borders to amount in words and tax summary sections
    sheet.getRange("A23:G25").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange("H23:K25").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

  } catch (error) {
    console.error('Error creating tax summary section:', error);
  }
}

function CreateBottomSection(sheet) {
  try {
    // Row 26: Terms and Conditions Header
    sheet.getRange("A26:G26").merge().setValue("Terms and Conditions")
         .setFontWeight("bold").setFontSize(13)
         .setHorizontalAlignment("center").setBackground("#FFFF00");
    sheet.setRowHeight(26, 25);

    // Rows 27-30: Terms and conditions content (merged across 4 rows)
    const termsText = "1. This is an electronically generated invoice.\n" +
                     "2. All disputes are subject to GUDUR jurisdiction only.\n" +
                     "3. If the Consignee makes any Inter State Sales, he has to pay GST himself.\n" +
                     "4. Goods once sold cannot be taken back or exchanged.\n" +
                     "5. Payment terms: As per agreement between buyer and seller.";

    sheet.getRange("A27:G30").merge().setValue(termsText)
         .setFontSize(11).setHorizontalAlignment("left")
         .setVerticalAlignment("top").setBackground("#FFFFF5")
         .setWrap(true);

    for (let i = 27; i <= 30; i++) {
      sheet.setRowHeight(i, 25);
    }

    // Tax summary on the right (columns H-K, rows 26-30)
    // Row 26: IGST (highlighted)
    sheet.getRange("H26:J26").merge().setValue("IGST :")
         .setFontWeight("bold").setFontSize(11)
         .setHorizontalAlignment("left").setBackground("#FFFFC8").setFontColor("#1a1a1a");

    sheet.getRange("K26").setValue("")
         .setFontWeight("bold").setHorizontalAlignment("right")
         .setBackground("#FFFFC8");

    // Row 27: CESS
    sheet.getRange("H27:J27").merge().setValue("CESS :")
         .setFontWeight("bold").setFontSize(11)
         .setHorizontalAlignment("left").setBackground("#F5F5F5").setFontColor("#1a1a1a");

    sheet.getRange("K27").setValue("")
         .setFontWeight("bold").setHorizontalAlignment("right")
         .setBackground("#D8DEE9");

    // Row 28: Total Tax (highlighted)
    sheet.getRange("H28:J28").merge().setValue("Total Tax:")
         .setFontWeight("bold").setFontSize(11)
         .setHorizontalAlignment("left").setVerticalAlignment("middle")
         .setBackground("#F0F0F0").setFontColor("#1a1a1a");

    sheet.getRange("K28").setValue("")
         .setFontWeight("bold").setHorizontalAlignment("right")
         .setBackground("#F0F0F0");

    // Rows 29-30: Total Amount After Tax (merged across 2 rows)
    sheet.getRange("H29:J30").merge().setValue("Total Amount After Tax:")
         .setFontWeight("bold").setFontSize(11)
         .setHorizontalAlignment("left").setVerticalAlignment("middle")
         .setBackground("#FFFFB4").setFontColor("#1a1a1a");

    sheet.getRange("K29:K30").merge().setValue("")
         .setFontWeight("bold").setFontSize(11)
         .setHorizontalAlignment("right").setVerticalAlignment("middle")
         .setBackground("#FFFFB4");

    sheet.setRowHeight(29, 18);
    sheet.setRowHeight(30, 18);

    // Apply borders to terms and tax summary sections
    sheet.getRange("A26:G30").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange("H26:K30").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

    // Create signature section (rows 31-37) - exactly matching Excel VBA
    CreateSignatureSection(sheet);

  } catch (error) {
    console.error('Error creating bottom section:', error);
  }
}

function CreateSignatureSection(sheet) {
  try {
    // Row 31: Signature headers with merged cells
    sheet.getRange("A31:C31").merge().setValue("Transporter")
         .setFontWeight("bold").setHorizontalAlignment("center")
         .setBackground("#DCDCDC");

    sheet.getRange("D31:G31").merge().setValue("Receiver")
         .setFontWeight("bold").setHorizontalAlignment("center")
         .setBackground("#DCDCDC");

    sheet.getRange("H31:K31").merge().setValue("Certified that the particulars given above are true and correct")
         .setFontWeight("bold").setFontSize(10)
         .setHorizontalAlignment("center").setVerticalAlignment("middle")
         .setWrap(true).setBackground("#DCDCDC");

    // Rows 32-33: Mobile Number Section (merged across 2 rows)
    sheet.getRange("A32:C33").merge().setValue("Mobile No: ___________________")
         .setFontSize(10).setHorizontalAlignment("center")
         .setVerticalAlignment("middle").setBackground("#FAFAFA");

    sheet.getRange("D32:G33").merge().setValue("Mobile No: ___________________")
         .setFontSize(10).setHorizontalAlignment("center")
         .setVerticalAlignment("middle").setBackground("#FAFAFA");

    sheet.getRange("H32:K33").merge().setValue("Mobile No: ___________________")
         .setFontSize(10).setHorizontalAlignment("center")
         .setVerticalAlignment("middle").setBackground("#FAFAFA");

    // Rows 34-36: Signature Space Section (merged across 3 rows)
    sheet.getRange("A34:C36").merge().setValue("")
         .setBackground("#FAFAFA");

    sheet.getRange("D34:G36").merge().setValue("")
         .setBackground("#FAFAFA");

    sheet.getRange("H34:K36").merge().setValue("")
         .setBackground("#FAFAFA");

    // Row 37: Signature Labels
    sheet.getRange("A37:C37").merge().setValue("Transporter's Signature")
         .setFontWeight("bold").setFontSize(10)
         .setHorizontalAlignment("center").setBackground("#D3D3D3");

    sheet.getRange("D37:G37").merge().setValue("Receiver's Signature")
         .setFontWeight("bold").setFontSize(10)
         .setHorizontalAlignment("center").setBackground("#D3D3D3");

    sheet.getRange("H37:K37").merge().setValue("Authorized Signatory")
         .setFontWeight("bold").setFontSize(10)
         .setHorizontalAlignment("center").setBackground("#D3D3D3");

    // Apply borders to entire signature section
    sheet.getRange("A31:K37").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

    // Set row heights
    sheet.setRowHeight(31, 35);
    for (let i = 32; i <= 36; i++) {
      sheet.setRowHeight(i, 25);
    }
    sheet.setRowHeight(37, 31);

  } catch (error) {
    console.error('Error creating signature section:', error);
  }
}

function AutoPopulateInvoiceFields(sheet) {
  try {
    // Auto-populate invoice number with next sequential number
    const nextInvoiceNumber = GetNextInvoiceNumber();
    sheet.getRange("C7").setValue(nextInvoiceNumber)
         .setFontWeight("bold").setFontColor("#DC143C")
         .setHorizontalAlignment("center").setVerticalAlignment("middle");

    // Set current date for Invoice Date and Date of Supply
    const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    sheet.getRange("C8").setValue(currentDate)
         .setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");

    sheet.getRange("F9").setValue(currentDate)
         .setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");

    sheet.getRange("G9").setValue(currentDate)
         .setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");

    // Reset state code to default
    sheet.getRange("C10").setValue("37")
         .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  } catch (error) {
    console.error('Error auto-populating invoice fields:', error);
  }
}

function SetupTaxCalculationFormulas(sheet) {
  try {
    // Set up formulas for row 18 (first item row) - exactly matching Excel VBA
    // Column G (Amount) = Quantity * Rate
    sheet.getRange("G18").setFormula('=IF(AND(D18<>"",F18<>""),D18*F18,"")');

    // Column H (Taxable Value) = Amount
    sheet.getRange("H18").setFormula('=IF(G18<>"",G18,"")');

    // Column I (IGST Rate) - VLOOKUP formula to get tax rate from HSN data
    sheet.getRange("I18").setFormula("=VLOOKUP(C18, warehouse!A:E, 5, FALSE)");

    // Column J (IGST Amount) = Taxable Value * IGST Rate / 100
    sheet.getRange("J18").setFormula('=IF(AND(H18<>"",I18<>""),H18*I18/100,"")');

    // Column K (Total Amount) = Taxable Value + IGST Amount
    sheet.getRange("K18").setFormula('=IF(AND(H18<>"",J18<>""),H18+J18,"")');

    // Format the formula cells
    sheet.getRange("G18:K18").setNumberFormat("0.00");
    sheet.getRange("I18").setNumberFormat("0.00");

    // Set up comprehensive tax summary formulas - exactly matching Excel VBA
    UpdateMultiItemTaxCalculations(sheet);

  } catch (error) {
    console.error('Error setting up tax calculation formulas:', error);
  }
}

function UpdateMultiItemTaxCalculations(sheet) {
  try {
    // Row 22: Total Quantity and Sub Total calculations - exactly matching Excel VBA
    sheet.getRange("D22").setFormula("=SUM(D18:D21)").setNumberFormat("#,##0.00");

    // Row 22: Sub Total calculations
    sheet.getRange("G22").setFormula("=SUM(G18:G21)"); // Amount column
    sheet.getRange("H22").setFormula("=SUM(H18:H21)"); // Taxable Value column
    sheet.getRange("G22:H22").setNumberFormat("#,##0.00");

    // Row 22: IGST and Total Amount
    sheet.getRange("I22").setFormula("=SUM(J18:J21)"); // IGST Amount column
    sheet.getRange("K22").setFormula("=SUM(K18:K21)"); // Total Amount column
    sheet.getRange("I22:K22").setNumberFormat("#,##0.00");

    // Row 23: Total Amount Before Tax
    sheet.getRange("K23").setFormula("=SUM(H18:H21)");

    // Row 24: CGST (0 for interstate)
    sheet.getRange("K24").setValue(0);

    // Row 25: SGST (0 for interstate)
    sheet.getRange("K25").setValue(0);

    // Row 26: IGST
    sheet.getRange("K26").setFormula("=SUM(J18:J21)");

    // Row 27: CESS (0 by default)
    sheet.getRange("K27").setValue(0);

    // Row 28: Total Tax
    sheet.getRange("K28").setFormula("=K24+K25+K26+K27");

    // Row 29-30: Total Amount After Tax
    sheet.getRange("K29").setFormula("=K23+K28");
    sheet.getRange("K30").setFormula("=K29");

    // Format all calculation cells
    sheet.getRange("K23:K30").setNumberFormat("#,##0.00");

    // Update Amount in Words (A24:G25 merged cell)
    sheet.getRange("A24").setFormula("=NumberToWords(K29)");

  } catch (error) {
    console.error('Error updating multi-item tax calculations:', error);
  }
}

function AutoFillConsigneeFromReceiver(sheet) {
  try {
    // Copy receiver details to consignee section
    const receiverName = sheet.getRange("C12").getValue();
    const receiverAddress = sheet.getRange("C13").getValue();
    const receiverGSTIN = sheet.getRange("C14").getValue();
    const receiverState = sheet.getRange("C15").getValue();
    const receiverStateCode = sheet.getRange("C16").getValue();

    if (receiverName) {
      sheet.getRange("I12").setValue(receiverName);
      sheet.getRange("I13").setValue(receiverAddress);
      sheet.getRange("I14").setValue(receiverGSTIN);
      sheet.getRange("I15").setValue(receiverState);
      sheet.getRange("I16").setValue(receiverStateCode);
    }
  } catch (error) {
    console.error('Error auto-filling consignee from receiver:', error);
  }
}

function AddNewItemRow(sheet) {
  try {
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GST_Tax_Invoice_for_interstate");
    }

    // Find the last item row (currently supports up to row 21)
    let lastItemRow = 21;

    // Check if we can add more rows (extend the table if needed)
    if (lastItemRow < 25) { // Allow up to 25 rows for items
      const newRow = lastItemRow + 1;
      const srNo = newRow - 17; // Calculate Sr.No.

      // Insert new row and copy formatting from previous row
      sheet.insertRowAfter(lastItemRow);

      // Set Sr.No.
      sheet.getRange(`A${newRow}`).setValue(srNo)
           .setHorizontalAlignment("center").setVerticalAlignment("middle");

      // Set formulas for calculated columns
      sheet.getRange(`G${newRow}`).setFormula(`=IF(AND(D${newRow}<>"",F${newRow}<>""),D${newRow}*F${newRow},"")`);
      sheet.getRange(`H${newRow}`).setFormula(`=G${newRow}`);
      sheet.getRange(`J${newRow}`).setFormula(`=IF(AND(H${newRow}<>"",I${newRow}<>""),H${newRow}*I${newRow}/100,"")`);
      sheet.getRange(`K${newRow}`).setFormula(`=IF(H${newRow}<>"",H${newRow}+J${newRow},"")`);

      // Apply formatting
      sheet.getRange(`A${newRow}:K${newRow}`).setBorder(true, true, true, true, false, false, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
      sheet.setRowHeight(newRow, 28);

      // Update tax summary formulas to include new row
      UpdateMultiItemTaxCalculations(sheet);

      SpreadsheetApp.getUi().alert('Success', 'New item row added successfully!', SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('Limit Reached', 'Maximum number of item rows reached.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error adding new item row: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
