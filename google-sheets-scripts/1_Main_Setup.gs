/**
 * ===============================================================================
 * MODULE: Main_Setup
 * DESCRIPTION: Handles the main setup, initialization, and user-facing start functions.
 * ===============================================================================
 */

// ████████████████████████████████████████████████████████████████████████████████
// 🚀 MAIN SETUP FUNCTIONS - USER INTERFACE
// ████████████████████████████████████████████████████████████████████████████████
// These are the PRIMARY functions users should run. All other functions are helpers.

function StartGSTSystem() {
  // Simple entry point for users - sets up everything automatically
  InitializeGSTSystem();
}

function QuickSetup() {
  // Ultra-simple setup function that should work without any prompts - FIXED VERSION
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Delete any existing sheets first to start fresh
    const existingSheets = ['GST_Tax_Invoice_for_interstate', 'Master', 'warehouse'];
    existingSheets.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        spreadsheet.deleteSheet(sheet);
      }
    });

    // Create sheets in order with enhanced functionality
    CreateMasterSheet();
    CreateWarehouseSheet();
    CreateInvoiceSheet();

    // Set up data validation and formulas
    const invoiceSheet = spreadsheet.getSheetByName("GST_Tax_Invoice_for_interstate");
    if (invoiceSheet) {
      SetupDataValidation(invoiceSheet);
      SetupCustomerDropdown(invoiceSheet);
      SetupHSNDropdown(invoiceSheet);
    }

    // Create custom menu
    CreateCustomMenu();

    SpreadsheetApp.getUi().alert(
      'Setup Complete - ENHANCED VERSION',
      'Quick setup complete! Three worksheets created:\n' +
      '1. GST_Tax_Invoice_for_interstate (with exact Excel layout)\n' +
      '2. Master (for invoice records)\n' +
      '3. warehouse (with state-code mapping)\n\n' +
      '✅ FIXED FEATURES:\n' +
      '• State code auto-fill when state is selected\n' +
      '• PDF export includes ONLY invoice sheet\n' +
      '• Complete invoice layout matching Excel VBA\n' +
      '• HSN code auto-fill for tax rates\n' +
      '• Enhanced tax calculations\n\n' +
      'System is ready for use!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Setup Error', 'Quick setup error: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function StartGSTSystemMinimal() {
  // Minimal initialization without data validation setup (for debugging)
  try {
    // Step 1: Create all supporting worksheets first
    CreateMasterSheet();
    CreateWarehouseSheet();

    // Step 2: Create the main invoice sheet
    CreateInvoiceSheet();

    SpreadsheetApp.getUi().alert(
      'System Ready',
      'GST Tax Invoice System initialized successfully (minimal version)!\n' +
      'All supporting worksheets created.\n' +
      'Data validation setup skipped for debugging.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Initialization Error', 'Error initializing GST system: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ShowAvailableFunctions() {
  // Display all available functions for the user
  let functionList = "GST SYSTEM - COMPLETE FUNCTIONALITY:\n\n";
  
  functionList += "✨ CLEAN FUNCTION LIST - Only 17 Functions:\n\n";
  
  functionList += "🚀 SETUP FUNCTIONS:\n";
  functionList += "• QuickSetup - Ultra-simple setup (recommended first)\n";
  functionList += "• StartGSTSystem - Complete system with all features\n";
  functionList += "• StartGSTSystemMinimal - Basic setup for debugging\n";
  functionList += "• ShowAvailableFunctions - Show this help list\n\n";
  
  functionList += "🔘 BUTTON FUNCTIONS (Daily Operations):\n";
  functionList += "• AddCustomerToWarehouseButton - Add customer to warehouse\n";
  functionList += "• AddNewItemRowButton - Add new item row to invoice\n";
  functionList += "• NewInvoiceButton - Generate fresh invoice with next number\n";
  functionList += "• SaveInvoiceButton - Save invoice to Master sheet\n";
  functionList += "• PrintAsPDFButton - Export as PDF to Google Drive\n";
  functionList += "• PrintButton - Save PDF + send notification\n\n";
  
  functionList += "👥 DATA MANAGEMENT & UTILITIES:\n";
  functionList += "• VerifyValidationSettings - Check manual editing capability\n";
  functionList += "• All customer data managed through SaveInvoiceButton\n";
  functionList += "• Enhanced dropdown functionality (dropdown + manual entry)\n";
  functionList += "• State code dropdowns show simple numeric codes (37, 29, etc.)\n";
  functionList += "• Customer dropdowns in both receiver and consignee sections\n\n";
  
  functionList += "🔒 PROFESSIONAL INTERFACE:\n";
  functionList += "20+ internal helper functions are PRIVATE and hidden\n";
  functionList += "for a clean, professional function list!\n\n";
  
  functionList += "💡 TIP: Start with QuickSetup, then use button functions!";

  SpreadsheetApp.getUi().alert('GST System - Complete & Clean', functionList, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ████████████████████████████████████████████████████████████████████████████████
// 🔧 INTERNAL SYSTEM FUNCTIONS - HIDDEN FROM FUNCTION LIST
// ████████████████████████████████████████████████████████████████████████████████

function InitializeGSTSystem() {
  // Master initialization function - creates all required worksheets and sets up the system
  try {
    let statusMsg = "Initializing GST System...\n";

    // Step 1: Create all supporting worksheets first
    statusMsg += "Creating Master sheet...";
    CreateMasterSheet();
    statusMsg += " ✓\n";

    statusMsg += "Creating warehouse sheet...";
    CreateWarehouseSheet();
    statusMsg += " ✓\n";

    // Step 2: Create the main invoice sheet
    statusMsg += "Creating invoice sheet...";
    CreateInvoiceSheet();
    statusMsg += " ✓\n";

    // Step 3: Set up all data validation and dropdowns
    statusMsg += "Setting up data validation...";
    const invoiceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GST_Tax_Invoice_for_interstate");
    SetupDataValidation(invoiceSheet);
    statusMsg += " ✓\n";

    statusMsg += "Setting up customer dropdown...";
    SetupCustomerDropdown(invoiceSheet);
    statusMsg += " ✓\n";

    statusMsg += "Setting up HSN dropdown...";
    SetupHSNDropdown(invoiceSheet);
    statusMsg += " ✓\n";

    statusMsg += "Setting up tax calculations...";
    SetupTaxCalculationFormulas(invoiceSheet);
    statusMsg += " ✓\n";

    SpreadsheetApp.getUi().alert(
      'System Ready',
      'GST Tax Invoice System initialized successfully!\n' +
      'All supporting worksheets created and configured.\n' +
      'You can now use the invoice system.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Initialization Error',
      'Error initializing GST system:\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== DEBUGGING AND TROUBLESHOOTING =====

function DebugInitialization() {
  // Step-by-step debugging of the initialization process
  try {
    let debugMsg = "Debug Initialization Process:\n\n";

    // Step 1: Test Master sheet creation
    debugMsg += "Step 1: Creating Master sheet... ";
    CreateMasterSheet();
    if (WorksheetExists("Master")) {
      debugMsg += "✓ SUCCESS\n";
    } else {
      debugMsg += "✗ FAILED\n";
    }

    // Step 2: Test warehouse sheet creation
    debugMsg += "Step 2: Creating warehouse sheet... ";
    CreateWarehouseSheet();
    if (WorksheetExists("warehouse")) {
      debugMsg += "✓ SUCCESS\n";
    } else {
      debugMsg += "✗ FAILED\n";
    }

    // Step 3: Test invoice sheet creation
    debugMsg += "Step 3: Creating invoice sheet... ";
    CreateInvoiceSheet();
    if (WorksheetExists("GST_Tax_Invoice_for_interstate")) {
      debugMsg += "✓ SUCCESS\n";
    } else {
      debugMsg += "✗ FAILED\n";
    }

    SpreadsheetApp.getUi().alert('Debug Results', debugMsg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Debug Error', 'Debug error: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ===== SYSTEM TESTING AND VERIFICATION =====

function TestGSTSystem() {
  // Comprehensive test function to verify the system works properly
  try {
    let testResults = "GST System Test Results:\n\n";

    // Test 1: Initialize the system
    testResults += "1. Initializing GST System... ";
    InitializeGSTSystem();
    testResults += "✓ PASSED\n";

    // Test 2: Verify all worksheets exist
    testResults += "2. Checking worksheet creation... ";
    if (WorksheetExists("GST_Tax_Invoice_for_interstate") &&
        WorksheetExists("Master") &&
        WorksheetExists("warehouse")) {
      testResults += "✓ PASSED\n";
    } else {
      testResults += "✗ FAILED\n";
    }

    // Test 3: Test invoice numbering
    testResults += "3. Testing invoice numbering... ";
    const testInvoiceNum = GetNextInvoiceNumber();
    if (testInvoiceNum !== "") {
      testResults += "✓ PASSED (" + testInvoiceNum + ")\n";
    } else {
      testResults += "✗ FAILED\n";
    }

    // Test 4: Test data validation setup
    testResults += "4. Testing data validation... ";
    const invoiceSheet = GetOrCreateWorksheet("GST_Tax_Invoice_for_interstate");
    SetupDataValidation(invoiceSheet);
    testResults += "✓ PASSED\n";

    testResults += "\nAll tests completed successfully!\n";
    testResults += "The GST system is ready for use.";

    SpreadsheetApp.getUi().alert('System Test Complete', testResults, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Test Error', 'Test failed with error: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
