# GST Tax Invoice System - Deployment Guide

This guide provides step-by-step instructions for deploying the Google Apps Script version of the GST Tax Invoice System.

## 📋 Prerequisites

Before starting the deployment, ensure you have:

- **Google Account** with access to Google Sheets and Google Drive
- **Google Drive folder** for PDF storage (or use the provided folder)
- **Basic understanding** of Google Sheets and Apps Script (helpful but not required)

## 🚀 Quick Deployment (Recommended)

### Step 1: Create Google Sheets Document
1. Open [Google Sheets](https://sheets.google.com)
2. Click **+ Blank** to create a new spreadsheet
3. Rename the spreadsheet to "GST Tax Invoice System"
   - Click on "Untitled spreadsheet" at the top
   - Type "GST Tax Invoice System"
   - Press Enter

### Step 2: Access Apps Script Editor
1. In your Google Sheet, click **Extensions** in the menu bar
2. Select **Apps Script** from the dropdown
3. A new tab will open with the Apps Script editor

### Step 3: Set Up the Project
1. **Rename the project**:
   - Click on "Untitled project" at the top
   - Type "GST Tax Invoice System"
   - Press Enter

2. **Delete default code**:
   - Select all content in the `Code.gs` file
   - Delete it (we'll replace it with our modules)

### Step 4: Import All Modules

#### Module 1: Main Setup
1. **Rename the file**:
   - Click on "Code.gs" in the left sidebar
   - Rename it to "1_Main_Setup"

2. **Copy the code**:
   - Open `1_Main_Setup.gs` from this repository
   - Copy all the content
   - Paste it into the Apps Script editor

#### Module 2-6: Remaining Modules
For each remaining module:

1. **Add new file**:
   - Click the **+** button next to "Files" in the left sidebar
   - Select **Script** from the dropdown

2. **Rename and add code**:
   - Rename the new file (e.g., "2_Module_InvoiceStructure")
   - Copy content from the corresponding `.gs` file
   - Paste into the editor

**Repeat for all modules**:
- `2_Module_InvoiceStructure`
- `3_Module_InvoiceEvents`
- `4_Module_Master`
- `5_Module_Warehouse`
- `6_Module_Utilities`

### Step 5: Enable Required Services
1. In the Apps Script editor, click **Services** in the left sidebar (+ icon)
2. Click **+ Add a service**
3. Find and add **Drive API**
4. Click **Add**

### Step 6: Configure Permissions
1. **Save the project**: Press `Ctrl+S` or click the save icon
2. **Run initial setup**:
   - Select `QuickSetup` from the function dropdown at the top
   - Click the **Run** button (▶️)

3. **Grant permissions**:
   - A dialog will appear asking for permissions
   - Click **Review permissions**
   - Select your Google account
   - Click **Allow** for each permission request:
     - View and manage your spreadsheets
     - View and manage your Google Drive files
     - Connect to an external service

### Step 7: Verify Installation
1. **Return to your Google Sheet** (switch back to the sheet tab)
2. **Refresh the page** (F5 or Ctrl+R)
3. **Check for custom menu**:
   - You should see "GST Invoice System" in the menu bar
   - If not visible, wait a few seconds and refresh again

4. **Test the system**:
   - Click **GST Invoice System** > **Setup** > **Quick Setup**
   - Verify that three sheets are created:
     - GST_Tax_Invoice_for_interstate
     - Master
     - warehouse

## 🔧 Advanced Configuration

### Customizing Company Details

1. **Open Apps Script editor**
2. **Navigate to** `2_Module_InvoiceStructure`
3. **Find lines 32-36** and update with your company information:

```javascript
CreateHeaderRow(sheet, 2, "A2:K2", "YOUR COMPANY NAME", 24, true, "#2F5061", "#FFFFFF", 37);
CreateHeaderRow(sheet, 3, "A3:K3", "Your Company Address", 11, true, "#F5F5F5", "#1a1a1a", 27);
CreateHeaderRow(sheet, 4, "A4:K4", "GSTIN: YOUR_GSTIN_NUMBER", 14, true, "#F5F5F5", "#1a1a1a", 27);
CreateHeaderRow(sheet, 5, "A5:K5", "Email: your-email@company.com", 11, true, "#F5F5F5", "#1a1a1a", 25);
```

### Configuring PDF Export Folder

#### Option 1: Use Provided Folder (Default)
The system is pre-configured to use the provided Google Drive folder. No changes needed.

#### Option 2: Use Your Own Folder
1. **Create a Google Drive folder** for storing PDFs
2. **Get the folder ID**:
   - Open the folder in Google Drive
   - Copy the folder ID from the URL (the long string after `/folders/`)
3. **Update the code**:
   - Open `3_Module_InvoiceEvents`
   - Find line 45: `const folderId = "1boyjaNQVZMZ6Gk_bRsTY7B0D7Lre_1r7";`
   - Replace with your folder ID: `const folderId = "YOUR_FOLDER_ID";`

### Adding Custom HSN Codes

1. **Open** `5_Module_Warehouse`
2. **Find the** `hsnData` array (around line 35)
3. **Add new entries** in the format:
```javascript
["HSN_CODE", "Description", CGST_Rate, SGST_Rate, IGST_Rate]
```

Example:
```javascript
["1234", "Your Product Description", 9, 9, 18]
```

## 🎯 Testing the Deployment

### Basic Functionality Test

1. **Create New Invoice**:
   - GST Invoice System > Invoice Operations > New Invoice
   - Verify invoice number is generated (INV-2024-001)
   - Check that current date is filled

2. **Fill Sample Data**:
   - Enter customer name: "Test Customer"
   - Enter item description: "Test Product"
   - Enter HSN code: "4401"
   - Enter quantity: 10
   - Enter rate: 100
   - Verify calculations work automatically

3. **Save Invoice**:
   - GST Invoice System > Invoice Operations > Save Invoice
   - Check that data appears in Master sheet

4. **Export PDF**:
   - GST Invoice System > Invoice Operations > Export as PDF
   - Verify PDF is created in Google Drive folder

### Validation Test

1. **Test Dropdowns**:
   - Click on customer name field (C12)
   - Verify dropdown appears with sample customers
   - Test that you can type custom values

2. **Test Calculations**:
   - Change quantity or rate in item rows
   - Verify amounts update automatically
   - Check tax calculations are correct

## 🛠️ Troubleshooting Deployment

### Common Issues and Solutions

#### 1. "Script function not found" Error
**Problem**: Functions not recognized
**Solution**: 
- Ensure all 6 modules are properly imported
- Check for typos in function names
- Save the project and try again

#### 2. Permission Denied Errors
**Problem**: Insufficient permissions
**Solution**:
- Re-run the setup function
- Grant all requested permissions
- Check Google Drive folder access

#### 3. Custom Menu Not Appearing
**Problem**: Menu doesn't show in Google Sheets
**Solution**:
- Refresh the Google Sheet page
- Wait 30 seconds after deployment
- Check that `onOpen()` function exists in the code

#### 4. PDF Export Fails
**Problem**: Cannot create PDF in Google Drive
**Solution**:
- Verify Google Drive folder permissions
- Check folder ID is correct
- Ensure Drive API is enabled

#### 5. Formulas Not Calculating
**Problem**: Tax calculations not working
**Solution**:
- Ensure all sheets are created properly
- Run `StartGSTSystem()` instead of `QuickSetup()`
- Check for formula syntax errors

### Getting Detailed Error Information

1. **Open Apps Script editor**
2. **Click on** "Executions" in the left sidebar
3. **View execution logs** for detailed error messages
4. **Use** `console.log()` statements for debugging

### Performance Optimization

1. **Reduce API calls** by batching operations
2. **Use** `SpreadsheetApp.flush()` sparingly
3. **Cache** frequently accessed data
4. **Minimize** cross-sheet references

## 📞 Support and Maintenance

### Regular Maintenance Tasks

1. **Monthly**: Check invoice numbering sequence
2. **Quarterly**: Backup Master sheet data
3. **Yearly**: Update HSN codes and tax rates
4. **As needed**: Add new customers to warehouse

### Backup Strategy

1. **Automatic**: Google Sheets provides automatic version history
2. **Manual**: Export Master sheet data to CSV monthly
3. **Drive**: PDFs are automatically stored in Google Drive

### Updates and Modifications

1. **Always test** changes in a copy first
2. **Document** any customizations made
3. **Keep** original code as backup
4. **Version control** using Apps Script's built-in versioning

## 🔒 Security Best Practices

1. **Limit sharing** of the Google Sheet to authorized users only
2. **Use** Google Workspace admin controls if available
3. **Regular review** of access permissions
4. **Monitor** execution logs for unusual activity
5. **Keep** sensitive data (like GSTIN) properly formatted

## 📈 Scaling Considerations

### For High Volume Usage

1. **Consider** using Google Sheets API for better performance
2. **Implement** data archiving for old invoices
3. **Use** separate sheets for different financial years
4. **Monitor** Google Apps Script quotas and limits

### Multi-User Environment

1. **Set up** proper sharing permissions
2. **Consider** using Google Workspace for better collaboration
3. **Implement** user-specific settings if needed
4. **Document** workflows for team members

This completes the deployment guide. The system should now be fully functional and ready for use!
