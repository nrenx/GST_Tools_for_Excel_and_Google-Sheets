# GST Tax Invoice System - File Structure

This document provides a complete overview of all files in the Google Apps Script version of the GST Tax Invoice System.

## 📁 Directory Structure

```
google-sheets-scripts/
├── 1_Main_Setup.gs              # Main system setup and initialization
├── 2_Module_InvoiceStructure.gs # Invoice sheet creation and formatting
├── 3_Module_InvoiceEvents.gs    # Button functions and user interactions
├── 4_Module_Master.gs           # Master sheet and invoice numbering
├── 5_Module_Warehouse.gs        # Data management and validation
├── 6_Module_Utilities.gs        # Shared helper functions
├── CONFIG.gs                    # Centralized configuration file
├── README.md                    # Complete system documentation
├── DEPLOYMENT_GUIDE.md          # Step-by-step deployment instructions
└── FILE_STRUCTURE.md           # This file - overview of all files
```

## 📋 Core Module Files

### 1. `1_Main_Setup.gs` - System Initialization
**Purpose**: Entry point for the GST system with setup and initialization functions.

**Key Functions**:
- `QuickSetup()` - Ultra-simple setup for first-time users
- `StartGSTSystem()` - Complete system setup with all features
- `StartGSTSystemMinimal()` - Basic setup for debugging
- `ShowAvailableFunctions()` - Display help and available functions
- `InitializeGSTSystem()` - Master initialization function
- `DebugInitialization()` - Step-by-step debugging
- `TestGSTSystem()` - Comprehensive system testing

**When to Use**: Run `QuickSetup()` first, then use other functions as needed.

### 2. `2_Module_InvoiceStructure.gs` - Invoice Layout
**Purpose**: Creates and formats the main GST invoice sheet with professional styling.

**Key Functions**:
- `CreateInvoiceSheet()` - Generate complete invoice layout
- `CreateHeaderRow()` - Create formatted header sections
- `CreateInvoiceDetailsSection()` - Invoice number, dates, transport details
- `CreatePartyDetailsSection()` - Receiver and consignee information
- `CreateItemTableSection()` - Product/service item table
- `CreateTaxSummarySection()` - Tax calculations and totals
- `CreateBottomSection()` - Bank details, declaration, signature
- `AutoPopulateInvoiceFields()` - Auto-fill invoice number and dates
- `SetupTaxCalculationFormulas()` - Configure automatic calculations
- `AddNewItemRow()` - Add additional item rows dynamically

**Features**:
- Professional muted slate blue styling (#2F5061)
- GST-compliant invoice format
- Automatic tax calculations
- Merged cells for clean layout
- Responsive column widths

### 3. `3_Module_InvoiceEvents.gs` - User Interactions
**Purpose**: Handles all button functions and user interactions for daily operations.

**Key Functions**:
- `NewInvoiceButton()` - Create new invoice with sequential numbering
- `SaveInvoiceButton()` - Save invoice to Master sheet for audit
- `PrintAsPDFButton()` - Export invoice as PDF to Google Drive
- `PrintButton()` - Save PDF and provide print instructions
- `AddCustomerToWarehouseButton()` - Add customer to warehouse data
- `AddNewItemRowButton()` - Add new item row to invoice
- `CreateCustomMenu()` - Create custom menu system
- `AutoFillConsigneeFromReceiverButton()` - Copy receiver to consignee
- `onOpen()` - Automatic menu creation when sheet opens
- `onEdit()` - Handle cell edit events

**Features**:
- Custom Google Sheets menu system
- PDF export to specified Google Drive folder
- Duplicate invoice detection and handling
- Automatic customer data validation
- Real-time calculation updates

### 4. `4_Module_Master.gs` - Data Management
**Purpose**: Manages the Master sheet for invoice records and automatic numbering.

**Key Functions**:
- `CreateMasterSheet()` - Generate GST-compliant audit sheet
- `GetNextInvoiceNumber()` - Generate sequential invoice numbers (INV-YYYY-NNN)
- `GetCurrentInvoiceNumber()` - Get the last used invoice number
- `ResetInvoiceCounter()` - Clear all invoice records (with confirmation)
- `GetInvoiceStats()` - Calculate invoice statistics and totals
- `ShowInvoiceStats()` - Display statistics to user

**Features**:
- GST audit-compliant record keeping
- Automatic invoice numbering by year
- Invoice statistics and reporting
- Data integrity checks
- Year-wise invoice tracking

### 5. `5_Module_Warehouse.gs` - Data Validation
**Purpose**: Creates warehouse sheet with master data and sets up dropdown validation.

**Key Functions**:
- `CreateWarehouseSheet()` - Generate warehouse with sample data
- `SetupDataValidation()` - Configure dropdown lists for invoice fields
- `SetupCustomerDropdown()` - Customer selection dropdowns
- `SetupHSNDropdown()` - HSN code dropdowns with tax rates

**Features**:
- Comprehensive HSN code database for wood products
- Indian states and state codes
- UOM (Unit of Measurement) options
- Transport mode options
- Customer master data management
- Flexible validation (allows manual entry)

### 6. `6_Module_Utilities.gs` - Helper Functions
**Purpose**: Shared utility functions used across all modules.

**Key Functions**:
- `NumberToWords()` - Convert amounts to words (Indian format)
- `WorksheetExists()` - Check if sheets exist
- `GetOrCreateWorksheet()` - Safely get or create sheets
- `CleanText()` - Text cleaning and normalization
- `VerifyValidationSettings()` - Display validation status
- `ValidateGSTIN()` - GSTIN format validation
- `ValidateEmail()` - Email format validation
- `ValidatePhoneNumber()` - Indian phone number validation
- `LogActivity()` - Activity logging for audit
- `GetSheetUrl()` - Get direct URLs to specific sheets

**Features**:
- Indian number-to-words conversion
- Comprehensive text cleaning
- Data validation helpers
- Error handling utilities
- Performance optimization functions

## 🔧 Configuration and Documentation Files

### 7. `CONFIG.gs` - Configuration Management
**Purpose**: Centralized configuration for easy customization.

**Configuration Sections**:
- **Company Information**: Name, address, GSTIN, bank details
- **Styling Configuration**: Colors, fonts, row heights, column widths
- **File Configuration**: Google Drive folder, sheet names, PDF naming
- **Invoice Configuration**: Numbering format, defaults, item table settings
- **Tax Configuration**: HSN codes and tax rates
- **Dropdown Configuration**: UOM, transport modes, states
- **System Configuration**: Validation settings, performance, feature flags

**Key Functions**:
- `getCompanyConfig()` - Access company information
- `getStyleConfig()` - Access styling settings
- `validateConfiguration()` - Validate all configuration values
- `testConfiguration()` - Test configuration integrity

### 8. `README.md` - Complete Documentation
**Purpose**: Comprehensive documentation for the entire system.

**Sections**:
- System overview and features
- Module structure and descriptions
- Setup instructions
- Usage guide for daily operations
- Customization instructions
- Troubleshooting guide
- Security and maintenance notes

### 9. `DEPLOYMENT_GUIDE.md` - Step-by-Step Setup
**Purpose**: Detailed deployment instructions for new installations.

**Sections**:
- Prerequisites and requirements
- Quick deployment steps
- Advanced configuration options
- Testing procedures
- Troubleshooting common issues
- Performance optimization
- Security best practices

### 10. `FILE_STRUCTURE.md` - This File
**Purpose**: Overview of all files and their relationships.

## 🔄 File Dependencies

### Import Order
When setting up the system, import files in this order:
1. `CONFIG.gs` - Configuration (optional but recommended)
2. `6_Module_Utilities.gs` - Utilities (required by all modules)
3. `4_Module_Master.gs` - Master sheet functions
4. `5_Module_Warehouse.gs` - Warehouse and validation
5. `2_Module_InvoiceStructure.gs` - Invoice creation
6. `3_Module_InvoiceEvents.gs` - User interactions
7. `1_Main_Setup.gs` - Main setup (calls all other modules)

### Function Dependencies
```
1_Main_Setup.gs
├── CreateMasterSheet() → 4_Module_Master.gs
├── CreateWarehouseSheet() → 5_Module_Warehouse.gs
├── CreateInvoiceSheet() → 2_Module_InvoiceStructure.gs
├── SetupDataValidation() → 5_Module_Warehouse.gs
└── WorksheetExists() → 6_Module_Utilities.gs

2_Module_InvoiceStructure.gs
├── GetNextInvoiceNumber() → 4_Module_Master.gs
├── CreateCustomMenu() → 3_Module_InvoiceEvents.gs
└── NumberToWords() → 6_Module_Utilities.gs

3_Module_InvoiceEvents.gs
├── GetNextInvoiceNumber() → 4_Module_Master.gs
├── AddNewItemRow() → 2_Module_InvoiceStructure.gs
└── CleanText() → 6_Module_Utilities.gs
```

## 📊 File Sizes and Complexity

| File | Lines | Functions | Complexity | Purpose |
|------|-------|-----------|------------|---------|
| 1_Main_Setup.gs | ~300 | 8 | Medium | System initialization |
| 2_Module_InvoiceStructure.gs | ~550 | 12 | High | Invoice layout creation |
| 3_Module_InvoiceEvents.gs | ~440 | 15 | High | User interactions |
| 4_Module_Master.gs | ~200 | 7 | Medium | Data management |
| 5_Module_Warehouse.gs | ~300 | 4 | Medium | Validation setup |
| 6_Module_Utilities.gs | ~300 | 15 | Medium | Helper functions |
| CONFIG.gs | ~300 | 8 | Low | Configuration |

**Total**: ~2,390 lines of code across 69 functions

## 🎯 Usage Patterns

### For New Users
1. Start with `README.md` for overview
2. Follow `DEPLOYMENT_GUIDE.md` for setup
3. Run `QuickSetup()` from `1_Main_Setup.gs`
4. Use custom menu for daily operations

### For Customization
1. Modify `CONFIG.gs` for basic changes
2. Edit specific modules for advanced customization
3. Test changes with `testConfiguration()`
4. Use `VerifyValidationSettings()` to check setup

### For Maintenance
1. Use `ShowInvoiceStats()` for reporting
2. Check execution logs in Apps Script editor
3. Run `TestGSTSystem()` for system verification
4. Monitor Google Drive folder for PDF exports

## 🔒 Security Considerations

### File Access
- All files run with user's Google account permissions
- No external API calls except Google services
- Data remains within user's Google Workspace

### Sensitive Information
- Company GSTIN and bank details in `CONFIG.gs`
- Customer data in warehouse sheet
- Invoice records in Master sheet
- PDF exports in Google Drive

### Best Practices
- Limit sharing of the Google Sheet
- Regular backup of Master sheet data
- Monitor execution logs for unusual activity
- Keep configuration file secure

This file structure provides a complete, modular, and maintainable GST Tax Invoice System that replicates all Excel VBA functionality while leveraging Google Workspace capabilities.
