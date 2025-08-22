Option Explicit
' ===============================================================================
' MODULE: RefreshButton
' DESCRIPTION: Comprehensive refresh button function that handles all refresh operations
'              including Sale Type display, tax calculations, dropdowns, and more
' ===============================================================================

Public Sub RefreshButton()
    ' Comprehensive refresh button function that handles all refresh operations
    Dim ws As Worksheet
    Dim saleType As String
    Dim refreshCount As Integer
    Dim refreshResults As String
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    refreshCount = 0
    refreshResults = "REFRESH OPERATIONS COMPLETED:" & vbCrLf & vbCrLf
    
    Set ws = ThisWorkbook.Worksheets("GST_Tax_Invoice_for_interstate")
    
    ' 1. Refresh Sale Type Display
    saleType = Trim(ws.Range("N7").Value)
    If saleType = "Interstate" Or saleType = "Intrastate" Then
        Call UpdateTaxFieldsDisplay(ws, saleType)
        refreshResults = refreshResults & "✅ Sale Type (" & saleType & ") tax fields updated" & vbCrLf
        refreshCount = refreshCount + 1
    Else
        refreshResults = refreshResults & "⚠️ Sale Type: Please select Interstate or Intrastate in N7" & vbCrLf
    End If
    
    ' 2. Refresh Tax Calculations
    Call UpdateMultiItemTaxCalculations(ws)
    refreshResults = refreshResults & "✅ Tax calculations refreshed" & vbCrLf
    refreshCount = refreshCount + 1
    
    ' 3. Refresh Data Validation Dropdowns
    Call SetupDataValidation(ws)
    refreshResults = refreshResults & "✅ Data validation dropdowns refreshed" & vbCrLf
    refreshCount = refreshCount + 1
    
    ' 4. Refresh Customer Auto-Population (if customer selected)
    Dim customerName As String
    customerName = Trim(ws.Range("C12").Value)
    If customerName <> "" Then
        Call SetupCustomerDropdown(ws)
        refreshResults = refreshResults & "✅ Customer dropdown refreshed" & vbCrLf
        refreshCount = refreshCount + 1
    End If
    
    ' 5. Refresh HSN Code Dropdowns
    Call SetupHSNDropdown(ws)
    refreshResults = refreshResults & "✅ HSN code dropdowns refreshed" & vbCrLf
    refreshCount = refreshCount + 1
    
    ' 6. Force worksheet recalculation
    ws.Calculate
    Application.Calculate
    refreshResults = refreshResults & "✅ Worksheet calculations updated" & vbCrLf
    refreshCount = refreshCount + 1
    
    ' 7. Clean empty product rows
    Call CleanEmptyProductRows(ws)
    refreshResults = refreshResults & "✅ Empty product rows cleaned" & vbCrLf
    refreshCount = refreshCount + 1
    
    Application.ScreenUpdating = True
    
    refreshResults = refreshResults & vbCrLf & "Total operations: " & refreshCount & vbCrLf & vbCrLf
    refreshResults = refreshResults & "🎉 All systems refreshed successfully!"
    
    MsgBox refreshResults, vbInformation, "System Refresh Complete"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error during refresh: " & Err.Description, vbCritical, "Refresh Error"
End Sub
