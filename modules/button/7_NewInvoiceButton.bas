Option Explicit
' ===============================================================================
' MODULE: NewInvoiceButton
' DESCRIPTION: Button function to generate a fresh invoice with next sequential 
'              number and cleared fields
' ===============================================================================

Public Sub NewInvoiceButton()
    ' Button function: Generate a fresh invoice with next sequential number and cleared fields
    Dim ws As Worksheet
    Dim response As VbMsgBoxResult
    Dim nextInvoiceNumber As String
    Dim clearRow As Long
    Dim i As Long
    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets("GST_Tax_Invoice_for_interstate")

    ' Confirm creating new invoice
    response = MsgBox("Create a new invoice?" & vbCrLf & "All current data will be cleared and a new invoice number will be generated.", vbYesNo + vbQuestion, "Confirm New Invoice")
    If response = vbNo Then Exit Sub

    ' Generate next sequential invoice number
    nextInvoiceNumber = GetNextInvoiceNumber()

    ' Clear and set invoice number (C7) with new sequential number
    With ws.Range("C7")
        .Value = nextInvoiceNumber
        .Font.Bold = True
        .Font.Color = RGB(220, 20, 60)  ' Red color for user input
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Set current date for Invoice Date (C8) and Date of Supply (F9, G9)
    With ws.Range("C8")
        .Value = Format(Date, "dd/mm/yyyy")
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("F9")
        .Value = Format(Date, "dd/mm/yyyy")
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("G9")
        .Value = Format(Date, "dd/mm/yyyy")
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' Reset state code to default (C10)
    With ws.Range("C10")
        .Value = "37"  ' Fixed value for Andhra Pradesh
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Clear all customer details (handle merged cells properly)
    On Error Resume Next
    ' Clear individual cells to avoid merged cell issues
    ws.Range("C12:H15").ClearContents ' Clear Receiver details (expanded range), preserving formulas in row 16
    ws.Range("K12:O15").ClearContents ' Clear Consignee details (expanded range), preserving formulas in row 16
    ws.Range("F7").Value = "By Lorry"   ' Reset Transport Mode
    ws.Range("F8").Value = ""           ' Clear Vehicle Number
    ws.Range("F10").Value = ""          ' Clear Place of Supply
    ws.Range("N10").Value = ""          ' Clear E-Way Bill No.
    ws.Range("N7").Value = "Interstate" ' Reset Sale Type to default
    On Error GoTo ErrorHandler

    ' Clear item table data (rows 18-21, keep headers and formulas) - EXPANDED TO ALL COLUMNS
    ws.Range("A18:O21").ClearContents
    ' Reset first Sr.No.
    ws.Range("A18").Value = 1

    ' Clear tax summary section (handle merged cells properly)
    On Error Resume Next
    ' Clear individual tax summary cells to avoid merged cell issues
    ws.Range("K23").ClearContents  ' Total Before Tax
    ws.Range("K24").ClearContents  ' CGST
    ws.Range("K25").ClearContents  ' SGST
    ws.Range("K26").ClearContents  ' IGST
    ws.Range("K27").ClearContents  ' CESS
    ws.Range("K28").ClearContents  ' Total Tax
    ws.Range("K29").ClearContents  ' Total After Tax
    ws.Range("K30").ClearContents  ' Grand Total
    ' Clear amount in words
    ws.Range("A31:K31").ClearContents  ' Amount in words row
    On Error GoTo ErrorHandler

    ' Update tax calculations
    Call UpdateMultiItemTaxCalculations(ws)

    MsgBox "New invoice created successfully!" & vbCrLf & "Invoice Number: " & nextInvoiceNumber & vbCrLf & "Date: " & Format(Date, "dd/mm/yyyy"), vbInformation, "New Invoice Ready"

    ' Select customer name field for new entry
    ws.Range("C12").Select
    Exit Sub

ErrorHandler:
    MsgBox "Error creating new invoice: " & Err.Description, vbCritical, "Error"
End Sub
