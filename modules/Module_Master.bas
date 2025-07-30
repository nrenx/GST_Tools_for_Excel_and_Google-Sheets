Option Explicit
' ===============================================================================
' MODULE: Module_Master
' DESCRIPTION: Handles all operations related to the 'Master' sheet, including
'              invoice record management and the invoice numbering system.
' ===============================================================================

' ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
' 📋 MASTER SHEET & INVOICE COUNTER FUNCTIONS
' ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

Public Sub CreateMasterSheet()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Master")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Master"

    With ws
        ' ===== GST INVOICE RECORDS FOR AUDIT & RETURN FILING (A1:P1) =====
        ' GST-compliant headers for complete invoice records
        .Range("A1").Value = "Invoice_Number"
        .Range("B1").Value = "Invoice_Date"
        .Range("C1").Value = "Customer_Name"
        .Range("D1").Value = "Customer_GSTIN"
        .Range("E1").Value = "Customer_State"
        .Range("F1").Value = "Customer_State_Code"
        .Range("G1").Value = "Total_Taxable_Value"
        .Range("H1").Value = "IGST_Rate"
        .Range("I1").Value = "IGST_Amount"
        .Range("J1").Value = "Total_Tax_Amount"
        .Range("K1").Value = "Total_Invoice_Value"
        .Range("L1").Value = "HSN_Codes"
        .Range("M1").Value = "Item_Description"
        .Range("N1").Value = "Quantity"
        .Range("O1").Value = "UOM"
        .Range("P1").Value = "Date_Created"

        ' Format GST audit headers
        .Range("A1:P1").Font.Bold = True
        .Range("A1:P1").Interior.Color = RGB(47, 80, 97)
        .Range("A1:P1").Font.Color = RGB(255, 255, 255)
        .Range("A1:P1").HorizontalAlignment = xlCenter
        .Range("A1:P1").WrapText = True
        .Rows(1).RowHeight = 30

        ' Add borders to header
        .Range("A1:P1").Borders.LineStyle = xlContinuous
        .Range("A1:P1").Borders.Color = RGB(204, 204, 204)

        ' Auto-fit columns for better visibility
        .Columns.AutoFit

        ' Set specific column widths for GST data
        .Columns("A").ColumnWidth = 15  ' Invoice Number
        .Columns("B").ColumnWidth = 12  ' Invoice Date
        .Columns("C").ColumnWidth = 20  ' Customer Name
        .Columns("D").ColumnWidth = 18  ' Customer GSTIN
        .Columns("E").ColumnWidth = 15  ' Customer State
        .Columns("F").ColumnWidth = 12  ' State Code
        .Columns("G").ColumnWidth = 15  ' Taxable Value
        .Columns("H").ColumnWidth = 10  ' IGST Rate
        .Columns("I").ColumnWidth = 12  ' IGST Amount
        .Columns("J").ColumnWidth = 12  ' Total Tax
        .Columns("K").ColumnWidth = 15  ' Invoice Value
        .Columns("L").ColumnWidth = 15  ' HSN Codes
        .Columns("M").ColumnWidth = 25  ' Item Description
        .Columns("N").ColumnWidth = 10  ' Quantity
        .Columns("O").ColumnWidth = 8   ' UOM
        .Columns("P").ColumnWidth = 12  ' Date Created

    End With
End Sub

Public Function GetNextInvoiceNumber() As String
    Dim masterWs As Worksheet
    Dim currentYear As Integer
    Dim counter As Integer
    Dim newInvoiceNumber As String
    Dim lastRow As Long
    Dim i As Long
    Dim maxCounter As Integer
    Dim invoiceNum As String

    On Error GoTo ErrorHandler

    ' Ensure supporting worksheets exist
    Call EnsureAllSupportingWorksheetsExist

    ' Get or create Master sheet
    Set masterWs = GetOrCreateWorksheet("Master")

    currentYear = Year(Date)
    maxCounter = 0

    ' Find the highest counter for the current year by examining existing invoice records
    lastRow = masterWs.Cells(masterWs.Rows.Count, "A").End(xlUp).Row

    If lastRow > 1 Then ' If there are invoice records
        For i = 2 To lastRow ' Start from row 2 (after header)
            invoiceNum = Trim(masterWs.Cells(i, "A").Value)
            If invoiceNum <> "" And InStr(invoiceNum, "INV-" & currentYear & "-") = 1 Then
                ' Extract counter from invoice number (format: INV-YYYY-NNN)
                maxCounter = Application.WorksheetFunction.Max(maxCounter, Val(Right(invoiceNum, 3)))
            End If
        Next i
    End If

    ' Set next counter
    counter = maxCounter + 1

    ' Generate new invoice number
    newInvoiceNumber = "INV-" & currentYear & "-" & Format(counter, "000")

    GetNextInvoiceNumber = newInvoiceNumber
    Exit Function

ErrorHandler:
    GetNextInvoiceNumber = "INV-" & Year(Date) & "-001"
End Function

Public Function GetCurrentInvoiceNumber() As String
    Dim masterWs As Worksheet
    Dim lastRow As Long
    Dim currentYear As Integer
    Dim maxCounter As Integer
    Dim i As Long
    Dim invoiceNum As String

    On Error GoTo ErrorHandler

    ' Ensure supporting worksheets exist
    Call EnsureAllSupportingWorksheetsExist

    Set masterWs = GetOrCreateWorksheet("Master")
    currentYear = Year(Date)
    maxCounter = 0

    If masterWs Is Nothing Then
        GetCurrentInvoiceNumber = "INV-" & currentYear & "-001"
        Exit Function
    End If

    ' Find the highest counter for the current year
    lastRow = masterWs.Cells(masterWs.Rows.Count, "A").End(xlUp).Row

    If lastRow > 1 Then ' If there are invoice records
        For i = 2 To lastRow ' Start from row 2 (after header)
            invoiceNum = Trim(masterWs.Cells(i, "A").Value)
            If invoiceNum <> "" And InStr(invoiceNum, "INV-" & currentYear & "-") = 1 Then
                ' Extract counter from invoice number (format: INV-YYYY-NNN)
                maxCounter = Application.WorksheetFunction.Max(maxCounter, Val(Right(invoiceNum, 3)))
            End If
        Next i
    End If

    If maxCounter = 0 Then
        GetCurrentInvoiceNumber = "INV-" & currentYear & "-001"
    Else
        GetCurrentInvoiceNumber = "INV-" & currentYear & "-" & Format(maxCounter, "000")
    End If
    Exit Function

ErrorHandler:
    GetCurrentInvoiceNumber = "INV-" & Year(Date) & "-001"
End Function

Public Sub ResetInvoiceCounter()
    Dim response As VbMsgBoxResult
    Dim masterWs As Worksheet
    Dim lastRow As Long

    response = MsgBox("WARNING: This will clear all invoice records from the Master sheet!" & vbCrLf & vbCrLf & _
                     "The invoice counter is now based on existing records in the Master sheet." & vbCrLf & _
                     "To reset numbering, you would need to clear the Master sheet." & vbCrLf & vbCrLf & _
                     "Are you sure you want to proceed?", vbYesNo + vbCritical, "Reset Invoice Counter")

    If response = vbYes Then
        Set masterWs = GetOrCreateWorksheet("Master")

        If Not masterWs Is Nothing Then
            ' Clear all invoice records (keep only the header row)
            lastRow = masterWs.Cells(masterWs.Rows.Count, "A").End(xlUp).Row
            If lastRow > 1 Then
                masterWs.Range("A2:P" & lastRow).ClearContents
            End If
            MsgBox "All invoice records cleared! Next invoice will be INV-" & Year(Date) & "-001", vbInformation, "Reset Complete"
        Else
            MsgBox "Master sheet not found!", vbExclamation
        End If
    End If
End Sub