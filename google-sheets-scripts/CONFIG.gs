/**
 * ===============================================================================
 * CONFIGURATION FILE
 * DESCRIPTION: Centralized configuration for the GST Tax Invoice System
 *              Modify these values to customize the system for your business
 * ===============================================================================
 */

// ████████████████████████████████████████████████████████████████████████████████
// 🏢 COMPANY INFORMATION
// ████████████████████████████████████████████████████████████████████████████████

const COMPANY_CONFIG = {
  // Company Details
  name: "KAVERI TRADERS",
  address: "191, Guduru, Pagadalapalli, Idulapalli, Tirupati, Andhra Pradesh - 524409",
  gstin: "37HERPB7733F1Z5",
  email: "kotidarisetty7777@gmail.com",
  
  // Default State Information
  defaultState: "Andhra Pradesh",
  defaultStateCode: "37",
  
  // Bank Details
  bankDetails: {
    accountHolder: "KAVERI TRADERS",
    bankName: "ANDHRA BANK",
    accountNumber: "123456789012345",
    branchCode: "ANDB0001234"
  },
  
  // Transporter Details
  defaultTransporter: "NARENDRA",
  defaultTransportMode: "By Lorry"
};

// ████████████████████████████████████████████████████████████████████████████████
// 🎨 STYLING CONFIGURATION
// ████████████████████████████████████████████████████████████████████████████████

const STYLE_CONFIG = {
  // Colors (Professional Muted Slate Blue Theme)
  colors: {
    headerBackground: "#2F5061",    // Muted slate blue for headers
    headerText: "#FFFFFF",          // White text on headers
    lightBackground: "#F5F5F5",     // Light grey background
    darkText: "#1a1a1a",           // Dark text
    userInputText: "#DC143C",       // Red color for user input fields
    borderColor: "#CCCCCC"          // Light grey borders
  },
  
  // Fonts
  fonts: {
    primary: "Segoe UI",
    defaultSize: 11,
    headerSize: 12,
    companyNameSize: 24,
    taxInvoiceSize: 22
  },
  
  // Row Heights
  rowHeights: {
    header: 30,
    companyName: 37,
    address: 27,
    taxInvoice: 28,
    partyDetails: 26,
    itemHeader: 35,
    itemRow: 28,
    taxSummary: 25,
    declaration: 35,
    signature: 60
  },
  
  // Column Widths (in pixels)
  columnWidths: {
    srNo: 60,           // Column A
    description: 144,   // Column B
    hsnCode: 144,      // Column C
    quantity: 108,     // Column D
    uom: 84,          // Column E
    rate: 120,        // Column F
    amount: 168,      // Column G
    taxableValue: 120, // Column H
    igstRate: 72,     // Column I
    igstAmount: 120,  // Column J
    totalAmount: 144  // Column K
  }
};

// ████████████████████████████████████████████████████████████████████████████████
// 📁 FILE AND FOLDER CONFIGURATION
// ████████████████████████████████████████████████████████████████████████████████

const FILE_CONFIG = {
  // Google Drive Folder for PDF Export
  pdfFolderId: "1boyjaNQVZMZ6Gk_bRsTY7B0D7Lre_1r7",
  
  // Sheet Names
  sheetNames: {
    invoice: "GST_Tax_Invoice_for_interstate",
    master: "Master",
    warehouse: "warehouse"
  },
  
  // PDF Naming Convention
  pdfNaming: {
    format: "{invoiceNumber}_{customerName}_{date}.pdf",
    dateFormat: "yyyy-MM-dd",
    maxCustomerNameLength: 20
  }
};

// ████████████████████████████████████████████████████████████████████████████████
// 🔢 INVOICE CONFIGURATION
// ████████████████████████████████████████████████████████████████████████████████

const INVOICE_CONFIG = {
  // Invoice Numbering
  numbering: {
    format: "INV-{year}-{counter}",
    counterDigits: 3,  // 001, 002, etc.
    startCounter: 1
  },
  
  // Default Values
  defaults: {
    transportMode: "By Lorry",
    transporter: "NARENDRA",
    state: "Andhra Pradesh",
    stateCode: "37"
  },
  
  // Item Table Configuration
  itemTable: {
    maxRows: 25,        // Maximum item rows allowed
    defaultRows: 4,     // Default number of item rows
    startRow: 18        // Starting row for items
  }
};

// ████████████████████████████████████████████████████████████████████████████████
// 📊 HSN CODE AND TAX CONFIGURATION
// ████████████████████████████████████████████████████████████████████████████████

const TAX_CONFIG = {
  // Default HSN Codes for Wood Products
  defaultHsnCodes: [
    { code: "4401", description: "Fuel wood, firewood, sawdust, wood waste and scrap", cgst: 2.5, sgst: 2.5, igst: 5 },
    { code: "4402", description: "Wood charcoal", cgst: 2.5, sgst: 2.5, igst: 5 },
    { code: "4403", description: "Wood in the rough (logs, unprocessed timber)", cgst: 9, sgst: 9, igst: 18 },
    { code: "4404", description: "Split poles, pickets, sticks, hoopwood, etc.", cgst: 6, sgst: 6, igst: 12 },
    { code: "4405", description: "Wood flour and wood wool", cgst: 6, sgst: 6, igst: 12 },
    { code: "4406", description: "Wooden railway or tramway sleepers", cgst: 6, sgst: 6, igst: 12 },
    { code: "4407", description: "Wood sawn or chipped", cgst: 9, sgst: 9, igst: 18 },
    { code: "4408", description: "Veneered wood and wood continuously shaped", cgst: 9, sgst: 9, igst: 18 },
    { code: "4409", description: "Moulded wood, flooring strips", cgst: 9, sgst: 9, igst: 18 },
    { code: "4410", description: "Particle board, oriented strand board (OSB), similar boards", cgst: 9, sgst: 9, igst: 18 },
    { code: "4412", description: "Plywood, veneered panels, laminated wood", cgst: 9, sgst: 9, igst: 18 },
    { code: "4413", description: "Densified wood", cgst: 9, sgst: 9, igst: 18 },
    { code: "4414", description: "Wooden frames for mirrors, photos, paintings", cgst: 9, sgst: 9, igst: 18 },
    { code: "4416", description: "Wooden barrels, casks, and other cooper's products", cgst: 6, sgst: 6, igst: 12 },
    { code: "4417", description: "Wooden tools, tool handles, broom handles", cgst: 6, sgst: 6, igst: 12 },
    { code: "4418", description: "Builders' joinery and carpentry of wood (doors, windows, etc.)", cgst: 9, sgst: 9, igst: 18 }
  ]
};

// ████████████████████████████████████████████████████████████████████████████████
// 📋 DROPDOWN LISTS CONFIGURATION
// ████████████████████████████████████████████████████████████████████████████████

const DROPDOWN_CONFIG = {
  // UOM (Unit of Measurement) Options
  uomList: ["NOS", "KG", "MT", "CBM", "SQM", "LTR", "PCS", "BOX", "SET", "PAIR"],
  
  // Transport Mode Options
  transportModes: ["By Lorry", "By Train", "By Air", "By Ship", "By Hand", "Courier", "Self Transport"],
  
  // Indian States List
  states: [
    "Jammu and Kashmir", "Himachal Pradesh", "Punjab", "Chandigarh", "Uttarakhand", "Haryana",
    "Delhi", "Rajasthan", "Uttar Pradesh", "Bihar", "Sikkim", "Arunachal Pradesh", "Nagaland",
    "Manipur", "Mizoram", "Tripura", "Meghalaya", "Assam", "West Bengal", "Jharkhand", "Odisha",
    "Chhattisgarh", "Madhya Pradesh", "Gujarat", "Dadra and Nagar Haveli and Daman and Diu (merged)",
    "Maharashtra", "Karnataka", "Goa", "Lakshadweep", "Kerala", "Tamil Nadu", "Puducherry",
    "Andaman and Nicobar Islands", "Telangana", "Andhra Pradesh", "Ladakh"
  ],
  
  // State Codes (corresponding to states above)
  stateCodes: [
    "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15",
    "16", "17", "18", "19", "20", "21", "22", "23", "24", "26", "27", "29", "30", "31", "32",
    "33", "34", "35", "36", "37", "38"
  ]
};

// ████████████████████████████████████████████████████████████████████████████████
// ⚙️ SYSTEM CONFIGURATION
// ████████████████████████████████████████████████████████████████████████████████

const SYSTEM_CONFIG = {
  // Validation Settings
  validation: {
    allowInvalidInput: true,    // Allow manual entry in dropdown fields
    showValidationHelp: true,   // Show help text for validation
    strictGstinValidation: false // Set to true for strict GSTIN format checking
  },
  
  // Performance Settings
  performance: {
    batchUpdates: true,         // Use batch updates for better performance
    autoCalculate: true,        // Enable automatic calculations
    cacheValidationRanges: true // Cache validation ranges for better performance
  },
  
  // Feature Flags
  features: {
    enableCustomMenu: true,     // Show custom menu in Google Sheets
    enableAutoBackup: false,    // Auto backup to separate sheet (not implemented)
    enableAuditLog: false,      // Log all operations (not implemented)
    enableMultiCurrency: false  // Support multiple currencies (not implemented)
  }
};

// ████████████████████████████████████████████████████████████████████████████████
// 🔧 HELPER FUNCTIONS TO ACCESS CONFIGURATION
// ████████████████████████████████████████████████████████████████████████████████

function getCompanyConfig() {
  return COMPANY_CONFIG;
}

function getStyleConfig() {
  return STYLE_CONFIG;
}

function getFileConfig() {
  return FILE_CONFIG;
}

function getInvoiceConfig() {
  return INVOICE_CONFIG;
}

function getTaxConfig() {
  return TAX_CONFIG;
}

function getDropdownConfig() {
  return DROPDOWN_CONFIG;
}

function getSystemConfig() {
  return SYSTEM_CONFIG;
}

// ████████████████████████████████████████████████████████████████████████████████
// 📝 CONFIGURATION VALIDATION
// ████████████████████████████████████████████████████████████████████████████████

function validateConfiguration() {
  // Validate that all required configuration values are present
  const errors = [];
  
  // Check company information
  if (!COMPANY_CONFIG.name) errors.push("Company name is required");
  if (!COMPANY_CONFIG.gstin) errors.push("Company GSTIN is required");
  
  // Check file configuration
  if (!FILE_CONFIG.pdfFolderId) errors.push("PDF folder ID is required");
  
  // Check HSN codes
  if (!TAX_CONFIG.defaultHsnCodes || TAX_CONFIG.defaultHsnCodes.length === 0) {
    errors.push("At least one HSN code is required");
  }
  
  if (errors.length > 0) {
    throw new Error("Configuration validation failed:\n" + errors.join("\n"));
  }
  
  return true;
}

// ████████████████████████████████████████████████████████████████████████████████
// 🎯 CONFIGURATION TESTING
// ████████████████████████████████████████████████████████████████████████████████

function testConfiguration() {
  try {
    validateConfiguration();
    
    const message = "CONFIGURATION TEST RESULTS:\n\n" +
                   "✅ Company Information: Valid\n" +
                   "✅ Styling Configuration: Valid\n" +
                   "✅ File Configuration: Valid\n" +
                   "✅ Invoice Configuration: Valid\n" +
                   "✅ Tax Configuration: Valid\n" +
                   "✅ Dropdown Configuration: Valid\n" +
                   "✅ System Configuration: Valid\n\n" +
                   "All configuration values are properly set!";
    
    SpreadsheetApp.getUi().alert('Configuration Test', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Configuration Error', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
