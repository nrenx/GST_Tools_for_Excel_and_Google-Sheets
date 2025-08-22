Option Explicit
' ===============================================================================
' MODULE: 15_DataPopulation
' DESCRIPTION: Handles data population for invoices including customer data,
'              HSN data, automatic field population, and data validation setup.
' ===============================================================================

' ████████████████████████████████████████████████████████████████████████████████
' 📝 DATA POPULATION FUNCTIONS
' ████████████████████████████████████████████████████████████████████████████████

Public Sub AutoPopulateInvoiceFields(ws As Worksheet)
    ' Auto-populate invoice number and dates with full manual override capability
    ' ALL auto-populated values can be manually edited by users
    Dim nextInvoiceNumber As String
    On Error GoTo ErrorHandler

    ' Auto-populate Invoice Number (Row 7, Column C)
    nextInvoiceNumber = GetNextInvoiceNumber()

    With ws.Range("C7")
        .Value = nextInvoiceNumber
        .Font.Bold = True
        .Font.Color = RGB(220, 20, 60)  ' Red color for user input
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        ' Allow manual editing - no validation restrictions
    End With

    ' Auto-populate Invoice Date (Row 8, Column C)
    With ws.Range("C8")
        .Value = Format(Date, "dd/mm/yyyy")
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        ' Allow manual editing - no validation restrictions
    End With

    ' Auto-populate Date of Supply (Row 9, Columns F & G)
    With ws.Range("F9")
        .Value = Format(Date, "dd/mm/yyyy")
        .HorizontalAlignment = xlLeft
        ' Allow manual editing - no validation restrictions
    End With

    With ws.Range("G9")
        .Value = Format(Date, "dd/mm/yyyy")
        .HorizontalAlignment = xlLeft
        ' Allow manual editing - no validation restrictions
    End With

    ' Set fixed State Code (Row 10, Column C) for Andhra Pradesh
    With ws.Range("C10")
        .Value = "37"
        .HorizontalAlignment = xlLeft
        ' No validation - fixed value
    End With

    Exit Sub

ErrorHandler:
    ' If auto-population fails, set default values
    ws.Range("C7").Value = "INV-" & Year(Date) & "-001"
    ws.Range("C8").Value = Format(Date, "dd/mm/yyyy")
    ws.Range("F9").Value = Format(Date, "dd/mm/yyyy")
    ws.Range("G9").Value = Format(Date, "dd/mm/yyyy")
End Sub

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
