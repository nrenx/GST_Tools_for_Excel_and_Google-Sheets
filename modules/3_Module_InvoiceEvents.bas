Option Explicit
' ===============================================================================
' MODULE: Module_InvoiceEvents  
' DESCRIPTION: Handles event handlers, data validation, and user interactions
'              on the invoice worksheet. Button functions are in separate modules.
' ===============================================================================

' ████████████████████████████████████████████████████████████████████████████████
' 📝 NOTE: BUTTON FUNCTIONS IN SEPARATE MODULES
' ████████████████████████████████████████████████████████████████████████████████
' Button functions are in separate .bas modules (import individually):
' • AddCustomerToWarehouseButton.bas - Customer warehouse management
' • NewInvoiceButton.bas - Fresh invoice generation
' • SaveInvoiceButton.bas - Invoice record saving
' • PrintAsPDFButton.bas - PDF export functionality
' • PrintButton.bas - Print operations
' • RefreshButton.bas - System refresh operations
' • ButtonManagement.bas - Button creation/removal
' • PDFUtilities.bas - PDF helper functions

' ████████████████████████████████████████████████████████████████████████████████
' 🔧 DATA VALIDATION & EVENT HANDLING
' ████████████████████████████████████████████████████████████████████████████████

Public Sub SetupDataValidation(ws As Worksheet)
    ' Setup data validation dropdowns for standardized inputs
    Dim validationWs As Worksheet
    On Error Resume Next

    ' Ensure supporting worksheets exist
    Call EnsureAllSupportingWorksheetsExist

    Set validationWs = GetOrCreateWorksheet("warehouse")

    If validationWs Is Nothing Then
        Exit Sub
    End If

    With ws
        ' UOM dropdown with manual text entry capability (Column E: 18-21)
        .Range("E18:E21").Validation.Delete
        .Range("E18:E21").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$G$2:$G$11"
        .Range("E18:E21").Validation.IgnoreBlank = True
        .Range("E18:E21").Validation.InCellDropdown = True
        .Range("E18:E21").Validation.ShowError = False

        ' Transport Mode dropdown with manual text entry capability (F7)
        .Range("F7").Validation.Delete
        .Range("F7").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$H$2:$H$8"
        .Range("F7").Validation.IgnoreBlank = True
        .Range("F7").Validation.InCellDropdown = True
        .Range("F7").Validation.ShowError = False

        ' State dropdown for Receiver (Row 15, Column C15:H15)
        .Range("C15").Validation.Delete
        .Range("C15").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$J$2:$J$37"
        .Range("C15").Validation.IgnoreBlank = True
        .Range("C15").Validation.InCellDropdown = True
        .Range("C15").Validation.ShowError = False

        ' State dropdown for Consignee (Row 15, Column K15:O15)
        .Range("K15").Validation.Delete
        .Range("K15").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$J$2:$J$37"
        .Range("K15").Validation.IgnoreBlank = True
        .Range("K15").Validation.InCellDropdown = True
        .Range("K15").Validation.ShowError = False

        ' Customer Name dropdown for Receiver (Row 12, Column C12:H12)
        .Range("C12").Validation.Delete
        .Range("C12").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$M$2:$M$50"
        .Range("C12").Validation.IgnoreBlank = True
        .Range("C12").Validation.InCellDropdown = True
        .Range("C12").Validation.ShowError = False

        ' Customer Name dropdown for Consignee (Row 12, Column K12:O12)
        .Range("K12").Validation.Delete
        .Range("K12").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$M$2:$M$50"
        .Range("K12").Validation.IgnoreBlank = True
        .Range("K12").Validation.InCellDropdown = True
        .Range("K12").Validation.ShowError = False

        ' GSTIN dropdown for Receiver (Row 14, Column C14:H14)
        .Range("C14").Validation.Delete
        .Range("C14").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$X$2:$X$50"
        .Range("C14").Validation.IgnoreBlank = True
        .Range("C14").Validation.InCellDropdown = True
        .Range("C14").Validation.ShowError = False

        ' GSTIN dropdown for Consignee (Row 14, Column K14:O14)
        .Range("K14").Validation.Delete
        .Range("K14").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$X$2:$X$50"
        .Range("K14").Validation.IgnoreBlank = True
        .Range("K14").Validation.InCellDropdown = True
        .Range("K14").Validation.ShowError = False
        
        ' Description dropdown for item (Row 18, Column B)
        .Range("B18").Validation.Delete
        .Range("B18").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$Z$2:$Z$10"
        .Range("B18").Validation.IgnoreBlank = True
        .Range("B18").Validation.InCellDropdown = True
        .Range("B18").Validation.ShowError = False

        ' Sale Type dropdown with manual text entry capability (N7:O7 merged)
        .Range("N7").Validation.Delete
        .Range("N7").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$AA$2:$AA$3"
        .Range("N7").Validation.IgnoreBlank = True
        .Range("N7").Validation.InCellDropdown = True
        .Range("N7").Validation.ShowError = False
    End With

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
