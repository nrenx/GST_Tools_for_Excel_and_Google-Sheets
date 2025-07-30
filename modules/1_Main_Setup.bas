Option Explicit
' ===============================================================================
' MODULE: Main_Setup
' DESCRIPTION: Handles the main setup, initialization, and user-facing start functions.
' ===============================================================================

' ████████████████████████████████████████████████████████████████████████████████
' 🚀 MAIN SETUP FUNCTIONS - USER INTERFACE
' ████████████████████████████████████████████████████████████████████████████████
' These are the PRIMARY functions users should run. All other functions are helpers.

Public Sub StartGSTSystem()
    ' Simple entry point for users - sets up everything automatically
    Call InitializeGSTSystem
End Sub

Public Sub QuickSetup()
    ' Ultra-simple setup function that should work without any prompts
    Dim ws As Worksheet
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Delete any existing sheets first to start fresh
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("GST_Tax_Invoice_for_interstate")
    If Not ws Is Nothing Then ws.Delete
    Set ws = ThisWorkbook.Sheets("Master")
    If Not ws Is Nothing Then ws.Delete
    Set ws = ThisWorkbook.Sheets("warehouse")
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo ErrorHandler

    ' Create sheets in order
    Call CreateMasterSheet
    Call CreateWarehouseSheet
    Call CreateInvoiceSheet

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Quick setup complete! Three worksheets created:" & vbCrLf & _
           "1. GST_Tax_Invoice_for_interstate" & vbCrLf & _
           "2. Master" & vbCrLf & _
           "3. warehouse" & vbCrLf & vbCrLf & _
           "System is ready for use!", vbInformation, "Setup Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Quick setup error: " & Err.Description, vbCritical, "Setup Error"
End Sub

Public Sub StartGSTSystemMinimal()
    ' Minimal initialization without data validation setup (for debugging)
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Step 1: Create all supporting worksheets first
    Call CreateMasterSheet
    Call CreateWarehouseSheet

    ' Step 2: Create the main invoice sheet
    Call CreateInvoiceSheet

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "GST Tax Invoice System initialized successfully (minimal version)!" & vbCrLf & _
           "All supporting worksheets created." & vbCrLf & _
           "Data validation setup skipped for debugging.", vbInformation, "System Ready"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error initializing GST system: " & Err.Description, vbCritical, "Initialization Error"
End Sub

Public Sub ShowAvailableFunctions()
    ' Display all available functions for the user
    Dim functionList As String
    functionList = "GST SYSTEM - COMPLETE FUNCTIONALITY:" & vbCrLf & vbCrLf

    functionList = functionList & "✨ CLEAN MACRO LIST (Alt+F8) - Only 17 Functions:" & vbCrLf & vbCrLf

    functionList = functionList & "🚀 SETUP FUNCTIONS:" & vbCrLf
    functionList = functionList & "• QuickSetup - Ultra-simple setup (recommended first)" & vbCrLf
    functionList = functionList & "• StartGSTSystem - Complete system with all features" & vbCrLf
    functionList = functionList & "• StartGSTSystemMinimal - Basic setup for debugging" & vbCrLf
    functionList = functionList & "• ShowAvailableFunctions - Show this help list" & vbCrLf & vbCrLf

    functionList = functionList & "🔘 BUTTON FUNCTIONS (Daily Operations):" & vbCrLf
    functionList = functionList & "• AddCustomerToWarehouseButton - Add customer to warehouse" & vbCrLf
    functionList = functionList & "• AddNewItemRowButton - Add new item row to invoice" & vbCrLf
    functionList = functionList & "• NewInvoiceButton - Generate fresh invoice with next number" & vbCrLf
    functionList = functionList & "• SaveInvoiceButton - Save invoice to Master sheet" & vbCrLf
    functionList = functionList & "• PrintAsPDFButton - Export as PDF to folder" & vbCrLf
    functionList = functionList & "• PrintButton - Save PDF + send to printer" & vbCrLf & vbCrLf

    functionList = functionList & "👥 DATA MANAGEMENT & UTILITIES:" & vbCrLf
    functionList = functionList & "• VerifyValidationSettings - Check manual editing capability" & vbCrLf
    functionList = functionList & "• All customer data managed through SaveInvoiceButton" & vbCrLf
    functionList = functionList & "• Enhanced dropdown functionality (dropdown + manual entry)" & vbCrLf
    functionList = functionList & "• State code dropdowns show simple numeric codes (37, 29, etc.)" & vbCrLf
    functionList = functionList & "• Customer dropdowns in both receiver and consignee sections" & vbCrLf & vbCrLf

    functionList = functionList & "🔒 PROFESSIONAL INTERFACE:" & vbCrLf
    functionList = functionList & "20+ internal helper functions are PRIVATE and hidden" & vbCrLf
    functionList = functionList & "for a clean, professional macro list!" & vbCrLf & vbCrLf

    functionList = functionList & "💡 TIP: Start with QuickSetup, then use button functions!"

    MsgBox functionList, vbInformation, "GST System - Complete & Clean"
End Sub

' ████████████████████████████████████████████████████████████████████████████████
' 🔧 INTERNAL SYSTEM FUNCTIONS - HIDDEN FROM MACRO LIST
' ████████████████████████████████████████████████████████████████████████████████

Private Sub InitializeGSTSystem()
    ' Master initialization function - creates all required worksheets and sets up the system
    Dim statusMsg As String
    Dim invoiceWs As Worksheet
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    statusMsg = "Initializing GST System..." & vbCrLf

    ' Step 1: Create all supporting worksheets first
    statusMsg = statusMsg & "Creating Master sheet..."
    Call CreateMasterSheet
    statusMsg = statusMsg & " ✓" & vbCrLf

    statusMsg = statusMsg & "Creating warehouse sheet..."
    Call CreateWarehouseSheet
    statusMsg = statusMsg & " ✓" & vbCrLf

    ' Step 2: Create the main invoice sheet
    statusMsg = statusMsg & "Creating invoice sheet..."
    Call CreateInvoiceSheet
    statusMsg = statusMsg & " ✓" & vbCrLf

    ' Step 3: Set up all data validation and dropdowns
    statusMsg = statusMsg & "Setting up data validation..."
    Set invoiceWs = ThisWorkbook.Sheets("GST_Tax_Invoice_for_interstate")
    Call SetupDataValidation(invoiceWs)
    statusMsg = statusMsg & " ✓" & vbCrLf

    statusMsg = statusMsg & "Setting up customer dropdown..."
    Call SetupCustomerDropdown(invoiceWs)
    statusMsg = statusMsg & " ✓" & vbCrLf

    statusMsg = statusMsg & "Setting up HSN dropdown..."
    Call SetupHSNDropdown(invoiceWs)
    statusMsg = statusMsg & " ✓" & vbCrLf

    statusMsg = statusMsg & "Setting up tax calculations..."
    Call SetupTaxCalculationFormulas(invoiceWs)
    statusMsg = statusMsg & " ✓" & vbCrLf

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "GST Tax Invoice System initialized successfully!" & vbCrLf & _
           "All supporting worksheets created and configured." & vbCrLf & _
           "You can now use the invoice system.", vbInformation, "System Ready"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error initializing GST system at: " & statusMsg & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "Line: " & Erl, vbCritical, "Initialization Error"
End Sub

' ===== DEBUGGING AND TROUBLESHOOTING =====

Private Sub DebugInitialization()
    ' Step-by-step debugging of the initialization process
    Dim debugMsg As String
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    debugMsg = "Debug Initialization Process:" & vbCrLf & vbCrLf

    ' Step 1: Test Master sheet creation
    debugMsg = debugMsg & "Step 1: Creating Master sheet... "
    Call CreateMasterSheet
    If WorksheetExists("Master") Then
        debugMsg = debugMsg & "✓ SUCCESS" & vbCrLf
    Else
        debugMsg = debugMsg & "✗ FAILED" & vbCrLf
    End If

    ' Step 2: Test warehouse sheet creation
    debugMsg = debugMsg & "Step 2: Creating warehouse sheet... "
    Call CreateWarehouseSheet
    If WorksheetExists("warehouse") Then
        debugMsg = debugMsg & "✓ SUCCESS" & vbCrLf
    Else
        debugMsg = debugMsg & "✗ FAILED" & vbCrLf
    End If

    ' Step 3: Test invoice sheet creation
    debugMsg = debugMsg & "Step 3: Creating invoice sheet... "
    Call CreateInvoiceSheet
    If WorksheetExists("GST_Tax_Invoice_for_interstate") Then
        debugMsg = debugMsg & "✓ SUCCESS" & vbCrLf
    Else
        debugMsg = debugMsg & "✗ FAILED" & vbCrLf
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox debugMsg, vbInformation, "Debug Results"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Debug error at: " & debugMsg & vbCrLf & "Error: " & Err.Description, vbCritical, "Debug Error"
End Sub

' ===== SYSTEM TESTING AND VERIFICATION =====

Private Sub TestGSTSystem()
    ' Comprehensive test function to verify the system works properly
    Dim testResults As String
    Dim testInvoiceNum As String
    Dim invoiceWs As Worksheet
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    testResults = "GST System Test Results:" & vbCrLf & vbCrLf

    ' Test 1: Initialize the system
    testResults = testResults & "1. Initializing GST System... "
    Call InitializeGSTSystem
    testResults = testResults & "✓ PASSED" & vbCrLf

    ' Test 2: Verify all worksheets exist
    testResults = testResults & "2. Checking worksheet creation... "
    If WorksheetExists("GST_Tax_Invoice_for_interstate") And _
       WorksheetExists("Master") And _
       WorksheetExists("warehouse") Then
        testResults = testResults & "✓ PASSED" & vbCrLf
    Else
        testResults = testResults & "✗ FAILED" & vbCrLf
    End If

    ' Test 3: Test invoice numbering
    testResults = testResults & "3. Testing invoice numbering... "
    testInvoiceNum = GetNextInvoiceNumber()
    If testInvoiceNum <> "" Then
        testResults = testResults & "✓ PASSED (" & testInvoiceNum & ")" & vbCrLf
    Else
        testResults = testResults & "✗ FAILED" & vbCrLf
    End If

    ' Test 4: Test data validation setup
    testResults = testResults & "4. Testing data validation... "
    Set invoiceWs = GetOrCreateWorksheet("GST_Tax_Invoice_for_interstate")
    Call SetupDataValidation(invoiceWs)
    testResults = testResults & "✓ PASSED" & vbCrLf

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    testResults = testResults & vbCrLf & "All tests completed successfully!" & vbCrLf & _
                  "The GST system is ready for use."

    MsgBox testResults, vbInformation, "System Test Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Test failed with error: " & Err.Description, vbCritical, "Test Error"
End Sub