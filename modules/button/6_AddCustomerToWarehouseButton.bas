Option Explicit
' ===============================================================================
' MODULE: AddCustomerToWarehouseButton
' DESCRIPTION: Button function to capture customer details from current invoice 
'              and save to warehouse for future use
' ===============================================================================

Public Sub AddCustomerToWarehouseButton()
    ' Button function: Capture customer details from current invoice and save to warehouse
    Dim invoiceWs As Worksheet
    Dim warehouseWs As Worksheet
    Dim customerName As String, address As String, gstin As String, stateCode As String
    Dim lastRow As Long
    Dim i As Long
    Dim newRow As Long
    On Error GoTo ErrorHandler

    ' Get worksheets
    Set invoiceWs = ThisWorkbook.Worksheets("GST_Tax_Invoice_for_interstate")
    Set warehouseWs = ThisWorkbook.Worksheets("warehouse")

    ' Get customer details from invoice
    customerName = Trim(invoiceWs.Range("C12").Value)
    address = Trim(invoiceWs.Range("C13").Value & " " & invoiceWs.Range("C14").Value & " " & invoiceWs.Range("C15").Value)
    gstin = Trim(invoiceWs.Range("C16").Value)
    stateCode = Trim(invoiceWs.Range("C10").Value)

    ' Validate required fields
    If customerName = "" Then
        MsgBox "Please enter customer name before adding to warehouse.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Check for duplicates in warehouse (Customer section - columns M-T)
    lastRow = warehouseWs.Cells(warehouseWs.Rows.Count, "M").End(xlUp).Row

    For i = 2 To lastRow ' Start from row 2 (skip header)
        If UCase(Trim(warehouseWs.Cells(i, "M").Value)) = UCase(customerName) Then
            MsgBox "Customer '" & customerName & "' already exists in warehouse.", vbInformation, "Duplicate Customer"
            Exit Sub
        End If
    Next i

    ' Add new customer to next available row
    newRow = lastRow + 1

    warehouseWs.Cells(newRow, "M").Value = customerName     ' Column M: Customer Name
    warehouseWs.Cells(newRow, "N").Value = address          ' Column N: Address
    warehouseWs.Cells(newRow, "O").Value = ""               ' Column O: State (empty for now)
    warehouseWs.Cells(newRow, "P").Value = stateCode        ' Column P: State Code
    warehouseWs.Cells(newRow, "Q").Value = gstin            ' Column Q: GSTIN
    warehouseWs.Cells(newRow, "R").Value = ""               ' Column R: Phone (empty)
    warehouseWs.Cells(newRow, "S").Value = ""               ' Column S: Email (empty)
    warehouseWs.Cells(newRow, "T").Value = ""               ' Column T: Contact Person (empty)

    MsgBox "Customer '" & customerName & "' added successfully to warehouse!", vbInformation, "Customer Added"
    Exit Sub

ErrorHandler:
    MsgBox "Error adding customer: " & Err.Description, vbCritical, "Error"
End Sub
