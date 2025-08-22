# GST Tax Invoice System - AI Coding Instructions

## Project Overview
This is a **dual-platform GST tax invoice system** supporting both **Excel VBA** and **Google Apps Script**. The system generates professional, GST-compliant invoices for Indian businesses with automatic tax calculations, sequential numbering (INV-YYYY-NNN), and audit trails.

## Architecture Patterns

### Modular Design Philosophy
The codebase follows a **6-module architecture** shared between VBA and Google Apps Script:
1. `1_Main_Setup` - Entry points and system initialization
2. `2_Module_InvoiceStructure` - Invoice layout generation and formatting  
3. `3_Module_InvoiceEvents` - User interactions and button handlers
4. `4_Module_Master` - Audit records and invoice numbering
5. `5_Module_Warehouse` - Data validation and customer management
6. `6_Module_Utilities` - Shared helper functions

### Cross-Platform Implementation
- **VBA modules**: `modules/*.bas` files for Excel
- **Google Apps Script**: `google-sheets-scripts/*.gs` files for Google Sheets
- **Shared configuration**: `CONFIG.gs` contains business-specific settings
- **Feature parity**: Both platforms implement identical GST compliance features

## Critical Workflows

### System Initialization
```vb
' Primary entry points (VBA)
QuickSetup()           ' Recommended first-time setup
StartGSTSystem()       ' Complete setup with data validation
StartGSTSystemMinimal() ' Debug-only minimal setup
```
The system auto-creates 3 worksheets: `GST_Tax_Invoice_for_interstate`, `Master`, `warehouse`

### Invoice Processing Flow
1. **Layout Creation** → `CreateInvoiceSheet()` builds professional formatted layout
2. **Auto-Population** → `AutoPopulateInvoiceFields()` sets invoice number/dates
3. **Tax Calculations** → `SetupTaxCalculationFormulas()` configures GST formulas
4. **Data Entry** → Users fill customer/item details via dropdowns + manual entry
5. **Audit Trail** → `SaveInvoiceButton()` stores complete records in Master sheet
6. **Export** → `PrintAsPDFButton()` generates PDFs in `invoices(demo)/` folder

### Tax Calculation Engine
The system handles **Interstate (IGST)** vs **Intrastate (CGST+SGST)** automatically:
- Columns I-J: IGST Rate/Amount (interstate only)
- Columns K-L: CGST Rate/Amount (intrastate only) 
- Columns M-N: SGST Rate/Amount (intrastate only)
- Column O: Total amount per line item

## Key Development Conventions

### Error Handling Pattern
```vb
On Error GoTo ErrorHandler
Application.ScreenUpdating = False
' ... main logic ...
Application.ScreenUpdating = True
Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
```

### Professional Styling Constants
- **Header color**: `#2F5061` (muted slate blue)
- **Background**: `#F5F5F5` (light grey for labels)
- **User input**: `#DC143C` (red text for user-editable fields)
- **Layout**: 16-column design (A-P) with specific column widths

### Data Validation Strategy
All dropdowns use `xlValidateList` with `IgnoreBlank:=True` to allow manual entry override. Source ranges are in the `warehouse` sheet for customer data, HSN codes, UOM options, etc.

## GST Compliance Requirements

### Master Sheet Structure
The Master sheet serves as the **GST audit trail** with mandatory fields:
- Invoice Number, Date, Customer GSTIN, State Codes
- Taxable Value, Tax Rates, Tax Amounts by type (IGST/CGST/SGST)
- HSN codes, Item descriptions, Total invoice value
- Creation timestamp for audit compliance

### Invoice Numbering System
Sequential numbering with `GetNextInvoiceNumber()`:
- Format: `INV-YYYY-NNN` (e.g., INV-2025-001)
- Year-based counter reset
- Duplicate prevention in Master sheet

### Tax Calculation Logic
- **Interstate**: Only IGST applicable (usually 12%)
- **Intrastate**: CGST + SGST (6% + 6% = 12% total)
- State detection drives tax type via customer state vs company state comparison

## Platform-Specific Notes

### Excel VBA Specifics
- Uses `ThisWorkbook.Sheets` for worksheet references
- Button creation via `ActiveSheet.Buttons.Add` with specific positioning
- PDF export uses `ExportAsFixedFormat` to designated file paths

### Google Apps Script Specifics  
- Uses `SpreadsheetApp.getActiveSpreadsheet()` for sheet access
- Menu system via `createMenu()` instead of buttons
- PDF export to Google Drive using `DriveApp.createFile()`
- Configuration in `CONFIG.gs` for business details

## Integration Points

### File Structure Dependencies
- `modules/` and `google-sheets-scripts/` maintain parallel functionality
- `CONFIG.gs` centralizes business configuration (company details, styling)
- `invoices(demo)/` contains sample PDF outputs
- `old-working-excel-modules/` preserves previous working versions

### External Dependencies
- **Excel**: Requires VBA editor access (`Alt+F11`)
- **Google Sheets**: Requires Apps Script permissions for Drive/Sheets API
- **PDF Generation**: Platform-specific export mechanisms
- **Data Persistence**: Worksheet-based storage (no external databases)

## Debugging Workflows

### System Verification
```vb
TestGSTSystem()        ' Verify all modules loaded correctly
ShowAvailableFunctions() ' Display help with available functions
VerifyValidationSettings() ' Check dropdown configurations
```

### Common Issue Patterns
- **Sheet naming**: Must use exact names (`GST_Tax_Invoice_for_interstate`)
- **Column references**: Fixed layout requires specific column letters
- **Formula dependencies**: Tax calculations depend on proper cell references
- **Cross-platform sync**: Changes require updates to both VBA and GAS versions

Focus on maintaining **GST compliance**, **professional formatting**, and **cross-platform feature parity** when making modifications.
[byterover-mcp]

# important 
always use byterover-retrieve-knowledge tool to get the related context before any tasks 
always use byterover-store-knowledge to store all the critical informations after sucessful tasks