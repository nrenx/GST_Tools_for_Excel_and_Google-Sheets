Option Explicit
' ===============================================================================
' MODULE: 16_WorksheetEventsSetup
' DESCRIPTION: Handles worksheet event setup, change monitoring, and event 
'              handling for sale type changes and state code management.
' ===============================================================================

' ████████████████████████████████████████████████████████████████████████████████
' 🔧 WORKSHEET EVENT HANDLING
' ████████████████████████████████████████████████████████████████████████████████

Public Sub SetupWorksheetChangeEvents(ws As Worksheet)
    ' Set up change monitoring for customer dropdown auto-population and Sale Type handling
    ' Since we're in a module, we'll ensure the worksheet change events are enabled
    On Error Resume Next

    ' Enable automatic calculation to ensure formulas update properly
    Application.Calculation = xlCalculationAutomatic
    
    ' Note: The actual worksheet change events for Sale Type are handled by 
    ' the HandleSaleTypeChange function in this module
    ' This function can be called manually when the Sale Type changes

    On Error GoTo 0
End Sub

Public Sub SetupStateCodeChangeEvents(ws As Worksheet)
    ' Simple state code setup - no automatic extraction needed
    ' State code dropdowns will show simple numeric codes only
    On Error GoTo 0
End Sub

' ████████████████████████████████████████████████████████████████████████████████
' 🎯 SALE TYPE EVENT HANDLING
' ████████████████████████████████████████████████████████████████████████████████

Public Sub HandleSaleTypeChange(ws As Worksheet, changedRange As Range)
    ' Handle Sale Type dropdown changes to update tax field display dynamically
    On Error Resume Next

    ' Check if the changed cell is the Sale Type dropdown (N7)
    If Not Intersect(changedRange, ws.Range("N7")) Is Nothing Then
        Dim saleType As String
        saleType = Trim(ws.Range("N7").Value)

        ' Validate sale type and update display
        If saleType = "Interstate" Or saleType = "Intrastate" Then
            Call UpdateTaxFieldsDisplay(ws, saleType)

            ' Recalculate the worksheet to update formulas
            ws.Calculate
        End If
    End If

    On Error GoTo 0
End Sub

Public Sub RefreshSaleTypeDisplay()
    ' Manual function to refresh Sale Type display - users can run this after changing Sale Type
    Dim ws As Worksheet
    Dim saleType As String
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("GST_Tax_Invoice_for_interstate")
    saleType = Trim(ws.Range("N7").Value)
    
    If saleType = "Interstate" Or saleType = "Intrastate" Then
        Call UpdateTaxFieldsDisplay(ws, saleType)
        ws.Calculate
        MsgBox "Tax fields updated for " & saleType & " sale type!", vbInformation, "Sale Type Updated"
    Else
        MsgBox "Please select either 'Interstate' or 'Intrastate' in cell N7.", vbExclamation, "Invalid Sale Type"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error updating sale type: " & Err.Description, vbCritical, "Error"
End Sub
