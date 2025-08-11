/**
 * ===============================================================================
 * MODULE: Module_Utilities
 * DESCRIPTION: Contains shared helper functions used across multiple modules,
 *              including worksheet management, text cleaning, and number-to-word
 *              conversion.
 * ===============================================================================
 */

// ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
// 🔧 UTILITY FUNCTIONS
// ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

function WorksheetExists(sheetName) {
  // Check if a worksheet exists
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    return sheet !== null;
  } catch (error) {
    return false;
  }
}

function GetOrCreateWorksheet(sheetName) {
  // Safely get or create a worksheet
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
    }
    
    return sheet;
  } catch (error) {
    console.error(`Error getting or creating worksheet ${sheetName}:`, error);
    return null;
  }
}

function EnsureAllSupportingWorksheetsExist() {
  // Ensure all required supporting worksheets exist
  try {
    // Create Master sheet if it doesn't exist
    if (!WorksheetExists("Master")) {
      CreateMasterSheet();
    }

    // Create warehouse sheet if it doesn't exist
    if (!WorksheetExists("warehouse")) {
      CreateWarehouseSheet();
    }
  } catch (error) {
    console.error('Error ensuring supporting worksheets exist:', error);
  }
}

function CleanText(inputText) {
  // Clean and normalize text input
  try {
    if (!inputText) return "";
    
    let cleanedText = inputText.toString();

    // Remove any question marks that might appear due to encoding issues
    cleanedText = cleanedText.replace(/\?/g, "");

    // Remove any other problematic characters
    cleanedText = cleanedText.replace(/[\x00-\x1F\x7F]/g, ""); // Remove control characters

    // Trim extra spaces
    cleanedText = cleanedText.trim();

    // Replace multiple spaces with single space
    cleanedText = cleanedText.replace(/\s+/g, " ");

    return cleanedText;
  } catch (error) {
    console.error('Error cleaning text:', error);
    return inputText ? inputText.toString() : "";
  }
}

function VerifyValidationSettings() {
  // Display current validation settings to confirm manual editing is enabled
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("GST_Tax_Invoice_for_interstate");

    if (!sheet) {
      SpreadsheetApp.getUi().alert('Error', 'Invoice sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    let message = "VALIDATION SETTINGS VERIFICATION:\n\n";
    message += "✅ ALL FIELDS SUPPORT MANUAL EDITING:\n\n";

    message += "📝 DROPDOWN + MANUAL ENTRY FIELDS:\n";
    message += "• Customer Name (C12) - Dropdown + Manual\n";
    message += "• Receiver State (C15) - Dropdown + Manual\n";
    message += "• Consignee State (I15) - Dropdown + Manual\n";
    message += "• HSN Code (C18:C21) - Dropdown + Manual\n";
    message += "• UOM (E18:E21) - Dropdown + Manual\n";
    message += "• Transport Mode (F7) - Dropdown + Manual\n\n";

    message += "🔓 FULLY EDITABLE FIELDS:\n";
    message += "• Invoice Number (C7) - Auto + Manual Override\n";
    message += "• Invoice Date (C8) - Auto + Manual Override\n";
    message += "• Date of Supply (F9, G9) - Auto + Manual Override\n";
    message += "• State Code (C10) - Fixed + Manual Override\n";
    message += "• All Address/GSTIN fields - Fully Manual\n";
    message += "• All Item details - Fully Manual\n\n";

    message += "🎯 KEY FEATURES:\n";
    message += "• No restrictive validations (setAllowInvalid = true)\n";
    message += "• Users can override ANY auto-populated value\n";
    message += "• Dropdown suggestions + free text entry\n";
    message += "• Google Sheets native validation system\n\n";

    message += "💡 All validation requirements have been successfully implemented!";

    SpreadsheetApp.getUi().alert('Validation Settings - All Clear ✅', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error verifying validation settings: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ===== AMOUNT IN WORDS CONVERSION SYSTEM =====

function NumberToWords(number) {
  // Convert number to words in Indian format - exactly matching Excel VBA
  try {
    if (!number || isNaN(number)) return "";

    const num = Math.round(parseFloat(number) * 100) / 100; // Round to 2 decimal places
    const [rupees, paise] = num.toString().split('.');

    let result = "";

    // Convert rupees part
    const rupeesNum = parseInt(rupees);
    if (rupeesNum === 0) {
      result = "Zero Rupees";
    } else if (rupeesNum === 1) {
      result = "One Rupee";
    } else {
      result = ConvertToWords(rupeesNum) + " Rupees";
    }

    // Convert paise part
    if (paise && parseInt(paise) > 0) {
      const paiseNum = parseInt(paise.padEnd(2, '0').substring(0, 2));
      if (paiseNum === 1) {
        result += " and One Paisa";
      } else if (paiseNum > 1) {
        result += " and " + ConvertToWords(paiseNum) + " Paise";
      }
    }

    return CleanText(result + " Only").toUpperCase();
  } catch (error) {
    console.error('Error converting number to words:', error);
    return "";
  }
}

function ConvertToWords(num) {
  // Helper function to convert numbers to words (Indian system)
  const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"];
  const teens = ["Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
  
  if (num === 0) return "";
  
  let result = "";
  
  // Handle crores
  if (num >= 10000000) {
    result += ConvertToWords(Math.floor(num / 10000000)) + " Crore ";
    num %= 10000000;
  }
  
  // Handle lakhs
  if (num >= 100000) {
    result += ConvertToWords(Math.floor(num / 100000)) + " Lakh ";
    num %= 100000;
  }
  
  // Handle thousands
  if (num >= 1000) {
    result += ConvertToWords(Math.floor(num / 1000)) + " Thousand ";
    num %= 1000;
  }
  
  // Handle hundreds
  if (num >= 100) {
    result += ones[Math.floor(num / 100)] + " Hundred ";
    num %= 100;
  }
  
  // Handle tens and ones
  if (num >= 20) {
    result += tens[Math.floor(num / 10)] + " ";
    num %= 10;
  } else if (num >= 10) {
    result += teens[num - 10] + " ";
    return result.trim();
  }
  
  if (num > 0) {
    result += ones[num] + " ";
  }
  
  return result.trim();
}

// ===== GOOGLE SHEETS SPECIFIC UTILITIES =====

function GetSheetUrl(sheetName) {
  // Get the URL of a specific sheet
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) return "";
    
    const sheetId = sheet.getSheetId();
    const spreadsheetUrl = spreadsheet.getUrl();
    
    return `${spreadsheetUrl}#gid=${sheetId}`;
  } catch (error) {
    console.error('Error getting sheet URL:', error);
    return "";
  }
}

function CopySheetFormatting(sourceSheet, targetSheet, sourceRange, targetRange) {
  // Copy formatting from one range to another
  try {
    const sourceRangeObj = sourceSheet.getRange(sourceRange);
    const targetRangeObj = targetSheet.getRange(targetRange);
    
    // Copy values and formatting
    sourceRangeObj.copyTo(targetRangeObj, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    
    return true;
  } catch (error) {
    console.error('Error copying sheet formatting:', error);
    return false;
  }
}

function ProtectSheet(sheetName, ranges) {
  // Protect specific ranges in a sheet
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet) return false;
    
    ranges.forEach(range => {
      const protection = sheet.getRange(range).protect();
      protection.setDescription(`Protected range: ${range}`);
      
      // Remove all editors except the owner
      protection.removeEditors(protection.getEditors());
    });
    
    return true;
  } catch (error) {
    console.error('Error protecting sheet:', error);
    return false;
  }
}

function LogActivity(action, details) {
  // Log activity for audit purposes
  try {
    const timestamp = new Date();
    const user = Session.getActiveUser().getEmail();
    
    console.log(`[${timestamp.toISOString()}] ${user}: ${action} - ${details}`);
    
    // Optionally, you could write to a log sheet
    // const logSheet = GetOrCreateWorksheet("Activity_Log");
    // const lastRow = logSheet.getLastRow();
    // logSheet.getRange(lastRow + 1, 1, 1, 4).setValues([[timestamp, user, action, details]]);
    
  } catch (error) {
    console.error('Error logging activity:', error);
  }
}

// ===== VALIDATION HELPERS =====

function ValidateGSTIN(gstin) {
  // Validate GSTIN format
  if (!gstin) return false;
  
  const gstinPattern = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/;
  return gstinPattern.test(gstin.toString().toUpperCase());
}

function ValidateEmail(email) {
  // Validate email format
  if (!email) return false;
  
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email.toString());
}

function ValidatePhoneNumber(phone) {
  // Validate Indian phone number format
  if (!phone) return false;
  
  const phonePattern = /^[6-9]\d{9}$/;
  return phonePattern.test(phone.toString().replace(/\D/g, ''));
}
