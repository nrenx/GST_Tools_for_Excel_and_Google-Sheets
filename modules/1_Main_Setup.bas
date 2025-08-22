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

' ████████████████████████████████████████████████████████████████████████████████
' 📋 HELP AND INFORMATION FUNCTIONS
' ████████████████████████████████████████████████████████████████████████████████

Public Sub ShowAvailableFunctions()
    ' Display all available functions for users
    Dim functionList As String
    
    functionList = "GST INVOICE SYSTEM - AVAILABLE FUNCTIONS:" & vbCrLf & vbCrLf
    
    functionList = functionList & "🚀 SETUP FUNCTIONS:" & vbCrLf
    functionList = functionList & "• StartGSTSystem - Complete system setup with all features" & vbCrLf
    functionList = functionList & "• ShowAvailableFunctions - Display this help list" & vbCrLf
    functionList = functionList & "• RefreshSaleTypeDisplay - Update tax fields after changing Sale Type" & vbCrLf & vbCrLf

    functionList = functionList & "🔘 BUTTON FUNCTIONS (Import individual .bas files):" & vbCrLf
    functionList = functionList & "• AddCustomerToWarehouseButton - Add customer to warehouse" & vbCrLf
    functionList = functionList & "• NewInvoiceButton - Generate fresh invoice with next number" & vbCrLf
    functionList = functionList & "• SaveInvoiceButton - Save invoice to Master sheet" & vbCrLf
    functionList = functionList & "• RefreshButton - 🔄 Refresh all systems" & vbCrLf
    functionList = functionList & "• PrintAsPDFButton - Export as PDF to folder" & vbCrLf
    functionList = functionList & "• PrintButton - Save PDF + send to printer" & vbCrLf
    functionList = functionList & "• CreateInvoiceButtons - Create all buttons on worksheet" & vbCrLf
    functionList = functionList & "• CreateDirectoryIfNotExists - Helper for PDF directory creation" & vbCrLf
    functionList = functionList & "• RefreshSaleTypeDisplay - Update tax fields after changing Sale Type" & vbCrLf & vbCrLf

    functionList = functionList & "📊 SYSTEM INFORMATION:" & vbCrLf
    functionList = functionList & "• System automatically creates 3 sheets: Invoice, Master, warehouse" & vbCrLf
    functionList = functionList & "• Invoice numbering: INV-YYYY-NNN format" & vbCrLf
    functionList = functionList & "• Professional styling with muted slate blue headers" & vbCrLf
    functionList = functionList & "• Dynamic tax calculation (Interstate/Intrastate)" & vbCrLf
    functionList = functionList & "• Sale Type dropdown in N7:O7 with conditional tax field clearing" & vbCrLf
    functionList = functionList & "• 🔄 Refresh All button handles all refresh operations automatically" & vbCrLf
    functionList = functionList & "• PDF export to: /Users/narendrachowdary/development/GST(excel)/invoices(demo)/" & vbCrLf & vbCrLf

    functionList = functionList & "🎯 QUICK START:" & vbCrLf
    functionList = functionList & "1. Import ALL .bas modules (including button modules)" & vbCrLf
    functionList = functionList & "2. Run 'StartGSTSystem' ONCE - this sets up everything automatically" & vbCrLf
    functionList = functionList & "3. Use buttons on invoice sheet for daily operations" & vbCrLf
    functionList = functionList & "4. Change Sale Type in N7 dropdown, then click 'Refresh All' button" & vbCrLf & vbCrLf
    
    functionList = functionList & "⚠️ IMPORTANT: Don't run individual button functions manually!" & vbCrLf
    functionList = functionList & "Button functions are for Excel buttons and integration only." & vbCrLf & vbCrLf

    functionList = functionList & "💡 TIP: Use the 🔄 Refresh All button after making any changes to update everything!"

    MsgBox functionList, vbInformation, "GST Invoice System - Help"
End Sub

' ████████████████████████████████████████████████████████████████████████████████
' 🔧 CORE SYSTEM FUNCTIONS
' ████████████████████████████████████████████████████████████████████████████████

Public Sub InitializeGSTSystem()
    ' Initialize the complete GST system with all components
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Create all required sheets
    Call CreateInvoiceSheet
    Call CreateMasterSheet
    Call CreateWarehouseSheet
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "System initialization failed: " & Err.Description, vbCritical, "Initialization Error"
End Sub

' ===== END OF PRODUCTION CODE =====
