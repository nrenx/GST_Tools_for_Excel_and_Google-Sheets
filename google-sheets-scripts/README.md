# GST Tax Invoice System - Google Apps Script Version

This directory contains the complete Google Apps Script version of the GST Tax Invoice System, converted from Excel VBA to work seamlessly with Google Sheets and Google Workspace.

## 📋 Overview

The Google Apps Script version maintains complete feature parity with the Excel VBA system while leveraging Google Workspace capabilities:

- **Professional GST-compliant invoice creation** with the same styling and layout
- **Automatic invoice numbering** (INV-YYYY-NNN format)
- **Customer and warehouse data management**
- **Master sheet for invoice records** and GST audit trails
- **PDF export to Google Drive** with organized file management
- **Data validation and dropdown functionality**
- **Tax calculations and GST compliance features**

## 🗂️ Module Structure

### 1. `1_Main_Setup.gs` (Main Module)
- **Purpose**: System initialization and user-facing start functions
- **Key Functions**:
  - `QuickSetup()` - Ultra-simple setup (recommended for first-time use)
  - `StartGSTSystem()` - Complete system setup with all features
  - `ShowAvailableFunctions()` - Display help and available functions

### 2. `2_Module_InvoiceStructure.gs`
- **Purpose**: Invoice sheet creation, formatting, and layout
- **Key Functions**:
  - `CreateInvoiceSheet()` - Generate complete invoice layout
  - `AutoPopulateInvoiceFields()` - Auto-fill invoice number and dates
  - `SetupTaxCalculationFormulas()` - Configure automatic tax calculations
  - `AddNewItemRow()` - Add additional item rows dynamically

### 3. `3_Module_InvoiceEvents.gs`
- **Purpose**: User interactions and button functions
- **Key Functions**:
  - `NewInvoiceButton()` - Create new invoice with next sequential number
  - `SaveInvoiceButton()` - Save invoice to Master sheet
  - `PrintAsPDFButton()` - Export as PDF to Google Drive
  - `AddCustomerToWarehouseButton()` - Add customer to warehouse data

### 4. `4_Module_Master.gs`
- **Purpose**: Master sheet operations and invoice numbering
- **Key Functions**:
  - `CreateMasterSheet()` - Generate GST-compliant audit sheet
  - `GetNextInvoiceNumber()` - Generate sequential invoice numbers
  - `ShowInvoiceStats()` - Display invoice statistics and totals

### 5. `5_Module_Warehouse.gs`
- **Purpose**: Data management and validation setup
- **Key Functions**:
  - `CreateWarehouseSheet()` - Generate warehouse with sample data
  - `SetupDataValidation()` - Configure dropdown lists
  - `SetupCustomerDropdown()` - Customer selection dropdowns
  - `SetupHSNDropdown()` - HSN code dropdowns

### 6. `6_Module_Utilities.gs`
- **Purpose**: Shared helper functions and utilities
- **Key Functions**:
  - `NumberToWords()` - Convert amounts to words (Indian format)
  - `WorksheetExists()` - Check if sheets exist
  - `CleanText()` - Text cleaning and normalization
  - `ValidateGSTIN()` - GSTIN format validation

## 🚀 Setup Instructions

### Step 1: Create New Google Sheets Document
1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Name it "GST Tax Invoice System"

### Step 2: Open Apps Script Editor
1. In your Google Sheet, go to **Extensions** > **Apps Script**
2. Delete the default `Code.gs` file content

### Step 3: Import All Modules
1. **Create Module 1**: 
   - Rename `Code.gs` to `1_Main_Setup`
   - Copy and paste the content from `1_Main_Setup.gs`

2. **Create Remaining Modules**:
   - Click the **+** button to add new files
   - Create files for each module: `2_Module_InvoiceStructure`, `3_Module_InvoiceEvents`, etc.
   - Copy and paste the respective content from each `.gs` file

### Step 4: Configure Google Drive Access
1. In the Apps Script editor, go to **Services** (left sidebar)
2. Click **+ Add a service**
3. Select **Drive API** and click **Add**

### Step 5: Set Up Permissions
1. Click **Run** on any function (e.g., `QuickSetup`)
2. Grant necessary permissions when prompted:
   - Google Sheets access
   - Google Drive access
   - Email access (for user identification)

### Step 6: Configure PDF Export Folder
The system is pre-configured to save PDFs to the specified Google Drive folder:
- **Folder ID**: `1boyjaNQVZMZ6Gk_bRsTY7B0D7Lre_1r7`
- **Folder URL**: https://drive.google.com/drive/folders/1boyjaNQVZMZ6Gk_bRsTY7B0D7Lre_1r7?usp=drive_link

Ensure you have access to this folder, or modify the folder ID in `3_Module_InvoiceEvents.gs` line 45.

## 🎯 How to Use

### Initial Setup
1. **Run Quick Setup**: 
   - In Apps Script editor, select `QuickSetup` function
   - Click **Run** button
   - This creates all necessary sheets with sample data

2. **Access Custom Menu**:
   - Return to your Google Sheet
   - You'll see a new "GST Invoice System" menu in the menu bar
   - Use this menu for all invoice operations

### Daily Operations

#### Creating New Invoice
1. **GST Invoice System** > **Invoice Operations** > **New Invoice**
2. System automatically:
   - Generates next sequential invoice number
   - Sets current date
   - Clears previous data

#### Filling Invoice Details
1. **Customer Information**: Use dropdown or type manually
2. **Item Details**: Enter description, HSN code, quantity, rate
3. **Tax Calculations**: Automatic based on HSN code and rates

#### Saving and Exporting
1. **Save Invoice**: **GST Invoice System** > **Invoice Operations** > **Save Invoice**
2. **Export PDF**: **GST Invoice System** > **Invoice Operations** > **Export as PDF**
3. PDFs are automatically saved to the configured Google Drive folder

## 🔧 Customization

### Changing Company Details
Edit the company information in `2_Module_InvoiceStructure.gs`:
- Lines 32-36: Company name, address, GSTIN, email

### Modifying Styling
Update colors and fonts in the `CreateHeaderRow` function calls:
- **Header Color**: `#2F5061` (muted slate blue)
- **Background Color**: `#F5F5F5` (light grey)
- **Font**: "Segoe UI"

### Adding HSN Codes
Modify the `hsnData` array in `5_Module_Warehouse.gs` (lines 35-50) to add more HSN codes and tax rates.

### Changing PDF Export Folder
Update the `folderId` in `3_Module_InvoiceEvents.gs` line 45 with your desired Google Drive folder ID.

## 📊 Features

### ✅ Complete Feature Parity
- All Excel VBA functionality preserved
- Same professional styling and layout
- Identical invoice numbering system
- GST compliance maintained

### ✅ Google Workspace Integration
- Native Google Sheets formulas and functions
- Google Drive PDF export
- Real-time collaboration support
- Cloud-based data storage

### ✅ Enhanced Capabilities
- Custom menu system for easy access
- Automatic permission handling
- Better error handling and user feedback
- Mobile-friendly interface

## 🛠️ Troubleshooting

### Common Issues

1. **Permission Errors**:
   - Re-run setup functions to grant permissions
   - Check Google Drive folder access

2. **Menu Not Appearing**:
   - Refresh the Google Sheet
   - Ensure `onOpen()` function exists in the code

3. **PDF Export Fails**:
   - Verify Google Drive folder permissions
   - Check folder ID in the code

4. **Formulas Not Working**:
   - Ensure all sheets are created properly
   - Run `StartGSTSystem()` for complete setup

### Getting Help
- Check the Apps Script execution log for detailed error messages
- Use `ShowAvailableFunctions()` to see all available functions
- Verify all modules are properly imported

## 📝 Notes

- **Data Persistence**: All data is stored in Google Sheets, providing automatic cloud backup
- **Collaboration**: Multiple users can work on the system simultaneously
- **Version Control**: Google Sheets maintains automatic version history
- **Mobile Access**: System works on mobile devices through Google Sheets app
- **Integration**: Easy integration with other Google Workspace tools

## 🔒 Security

- All data remains in your Google Workspace
- Standard Google security and encryption apply
- Access controlled through Google account permissions
- Audit trail maintained in Master sheet
