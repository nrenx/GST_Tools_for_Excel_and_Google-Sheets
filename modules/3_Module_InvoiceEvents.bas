Option Explicit
' ===============================================================================
' MODULE: Module_InvoiceEvents
' DESCRIPTION: Handles all button clicks, event handlers, and user interactions
'              on the invoice worksheet.
' ===============================================================================

' ████████████████████████████████████████████████████████████████████████████████
' 🔘 BUTTON FUNCTIONS - DAILY OPERATIONS
' ████████████████████████████████████████████████████████████████████████████████
' These functions are designed to be assigned to Excel buttons for daily use.

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

Public Sub AddNewItemRowButton()
    ' Button function: Add new item row after existing item rows with clean layout
    Call AddNewItemRow
End Sub

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
    ws.Range("C12:F16").ClearContents ' Clear Receiver details
    ws.Range("I12:K16").ClearContents ' Clear Consignee details
    On Error GoTo ErrorHandler

    ' Clear item table data (rows 18-21, keep headers and formulas)
    ws.Range("A18:F21").ClearContents
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

Public Sub SaveInvoiceButton()
    ' Button function: Save complete invoice record to Master sheet for future reference
    Dim invoiceWs As Worksheet
    Dim masterWs As Worksheet
    Dim invoiceNumber As String, invoiceDate As String, customerName As String
    Dim customerGSTIN As String, customerState As String, customerStateCode As String
    Dim hsnCodes As String, itemDescriptions As String, totalQuantity As String, uomList As String
    Dim taxableTotal As Double, igstTotal As Double, grandTotal As Double, totalQty As Double
    Dim i As Long
    Dim igstRate As String
    Dim lastRow As Long
    Dim response As VbMsgBoxResult
    On Error GoTo ErrorHandler

    Set invoiceWs = ThisWorkbook.Worksheets("GST_Tax_Invoice_for_interstate")
    Set masterWs = ThisWorkbook.Worksheets("Master")

    ' Get invoice details for GST compliance
    invoiceNumber = Trim(invoiceWs.Range("C7").Value)
    invoiceDate = Trim(invoiceWs.Range("C8").Value)
    customerName = Trim(invoiceWs.Range("C12").Value)
    customerGSTIN = Trim(invoiceWs.Range("C14").Value)
    customerState = Trim(invoiceWs.Range("C15").Value)
    customerStateCode = Trim(invoiceWs.Range("C16").Value)

    ' Calculate totals and collect item details from item table
    For i = 18 To 21 ' Check all possible item rows
        If invoiceWs.Cells(i, "H").Value <> "" And IsNumeric(invoiceWs.Cells(i, "H").Value) Then
            taxableTotal = taxableTotal + invoiceWs.Cells(i, "H").Value
        End If
        If invoiceWs.Cells(i, "J").Value <> "" And IsNumeric(invoiceWs.Cells(i, "J").Value) Then
            igstTotal = igstTotal + invoiceWs.Cells(i, "J").Value
        End If
        If invoiceWs.Cells(i, "K").Value <> "" And IsNumeric(invoiceWs.Cells(i, "K").Value) Then
            grandTotal = grandTotal + invoiceWs.Cells(i, "K").Value
        End If

        ' Collect item details for GST audit
        If Trim(invoiceWs.Cells(i, "C").Value) <> "" Then ' HSN Code
            If hsnCodes <> "" Then hsnCodes = hsnCodes & "; "
            hsnCodes = hsnCodes & Trim(invoiceWs.Cells(i, "C").Value)
        End If
        If Trim(invoiceWs.Cells(i, "B").Value) <> "" Then ' Item Description
            If itemDescriptions <> "" Then itemDescriptions = itemDescriptions & "; "
            itemDescriptions = itemDescriptions & Trim(invoiceWs.Cells(i, "B").Value)
        End If
        If invoiceWs.Cells(i, "D").Value <> "" And IsNumeric(invoiceWs.Cells(i, "D").Value) Then ' Quantity
            totalQty = totalQty + invoiceWs.Cells(i, "D").Value
        End If
        If Trim(invoiceWs.Cells(i, "E").Value) <> "" Then ' UOM
            If uomList <> "" And InStr(uomList, Trim(invoiceWs.Cells(i, "E").Value)) = 0 Then
                uomList = uomList & "; "
            End If
            If InStr(uomList, Trim(invoiceWs.Cells(i, "E").Value)) = 0 Then
                uomList = uomList & Trim(invoiceWs.Cells(i, "E").Value)
            End If
        End If
    Next i

    ' Calculate IGST rate (assuming 18% for interstate)
    If taxableTotal > 0 Then
        igstRate = Format((igstTotal / taxableTotal) * 100, "0.00") & "%"
    Else
        igstRate = "18.00%"
    End If

    ' Validate required fields
    If invoiceNumber = "" Or customerName = "" Then
        MsgBox "Please ensure invoice number and customer name are filled before saving.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Check if invoice already exists in Master sheet (starting from row 2, after headers)
    lastRow = masterWs.Cells(masterWs.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 1 ' Ensure we start after the header row

    For i = 2 To lastRow ' Start from row 2 (skip header row)
        If Trim(masterWs.Cells(i, "A").Value) = invoiceNumber Then
            response = MsgBox("Invoice " & invoiceNumber & " already exists in Master sheet." & vbCrLf & "Update existing record?", vbYesNo + vbQuestion, "Duplicate Invoice")
            If response = vbNo Then Exit Sub
            ' Update existing record
            GoTo UpdateRecord
        End If
    Next i

    ' Add new record
    lastRow = lastRow + 1

UpdateRecord:
    ' Save complete GST-compliant invoice data to Master sheet
    With masterWs
        .Cells(lastRow, "A").Value = invoiceNumber          ' Column A: Invoice_Number
        .Cells(lastRow, "B").Value = invoiceDate            ' Column B: Invoice_Date
        .Cells(lastRow, "C").Value = customerName           ' Column C: Customer_Name
        .Cells(lastRow, "D").Value = customerGSTIN          ' Column D: Customer_GSTIN
        .Cells(lastRow, "E").Value = customerState          ' Column E: Customer_State
        .Cells(lastRow, "F").Value = customerStateCode      ' Column F: Customer_State_Code
        .Cells(lastRow, "G").Value = taxableTotal           ' Column G: Total_Taxable_Value
        .Cells(lastRow, "H").Value = igstRate               ' Column H: IGST_Rate
        .Cells(lastRow, "I").Value = igstTotal              ' Column I: IGST_Amount
        .Cells(lastRow, "J").Value = igstTotal              ' Column J: Total_Tax_Amount (same as IGST for interstate)
        .Cells(lastRow, "K").Value = grandTotal             ' Column K: Total_Invoice_Value
        .Cells(lastRow, "L").Value = hsnCodes               ' Column L: HSN_Codes
        .Cells(lastRow, "M").Value = itemDescriptions       ' Column M: Item_Description
        .Cells(lastRow, "N").Value = totalQty               ' Column N: Quantity
        .Cells(lastRow, "O").Value = uomList                ' Column O: UOM
        .Cells(lastRow, "P").Value = Now                    ' Column P: Date_Created

        ' Add borders for the new record
        .Range("A" & lastRow & ":P" & lastRow).Borders.LineStyle = xlContinuous
        .Range("A" & lastRow & ":P" & lastRow).Borders.Color = RGB(204, 204, 204)
    End With

    MsgBox "Invoice " & invoiceNumber & " saved successfully to Master sheet!" & vbCrLf & _
           "Customer: " & customerName & vbCrLf & _
           "Total Taxable Value: ₹" & Format(taxableTotal, "#,##0.00") & vbCrLf & _
           "IGST Amount: ₹" & Format(igstTotal, "#,##0.00") & vbCrLf & _
           "Total Invoice Value: ₹" & Format(grandTotal, "#,##0.00") & vbCrLf & vbCrLf & _
           "Record saved for GST audit and return filing purposes.", vbInformation, "GST Invoice Record Saved"
    Exit Sub

ErrorHandler:
    MsgBox "Error saving invoice: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub PrintAsPDFButton()
    ' Button function: Export invoice as a two-page PDF (Original and Duplicate)
    Dim originalWs As Worksheet
    Dim duplicateWs As Worksheet
    Dim invoiceNumber As String
    Dim cleanInvoiceNumber As String
    Dim pdfPath As String
    Dim fullPath As String
    On Error GoTo ErrorHandler

    Set originalWs = ThisWorkbook.Worksheets("GST_Tax_Invoice_for_interstate")

    ' Get invoice number for filename
    invoiceNumber = Trim(originalWs.Range("C7").Value)

    If invoiceNumber = "" Then
        MsgBox "Please ensure invoice number is filled before exporting to PDF.", vbExclamation, "Missing Invoice Number"
        Exit Sub
    End If

    ' Clean invoice number for filename
    cleanInvoiceNumber = Replace(Replace(Replace(invoiceNumber, "/", "-"), "\", "-"), ":", "-")

    ' Set PDF export path
    pdfPath = "/Users/narendrachowdary/development/GST(excel)/invoices(demo)/"

    ' Create directory if it doesn't exist
    On Error Resume Next
    If Dir(pdfPath, vbDirectory) = "" Then
        MkDir pdfPath
    End If
    On Error GoTo 0

    ' Full filename with path
    fullPath = pdfPath & cleanInvoiceNumber & ".pdf"

    ' Delete any existing temporary sheet to avoid errors
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("DuplicateInvoiceTemp").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create a temporary duplicate of the invoice sheet
    Application.DisplayAlerts = False
    originalWs.Copy After:=originalWs
    Set duplicateWs = ActiveSheet
    duplicateWs.Name = "DuplicateInvoiceTemp"
    Application.DisplayAlerts = True

    ' Change the header on the duplicate sheet
    duplicateWs.Range("A1").Value = "DUPLICATE"

    ' Set print area and page setup for the original sheet
    originalWs.PageSetup.PrintArea = "A1:K37"
    With originalWs.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .CenterHorizontally = True
        .CenterVertically = False
    End With

    ' Set print area and page setup for the duplicate sheet
    duplicateWs.PageSetup.PrintArea = "A1:K37"
    With duplicateWs.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .CenterHorizontally = True
        .CenterVertically = False
    End With

    ' Export both sheets to a single PDF
    ThisWorkbook.Sheets(Array(originalWs.Name, duplicateWs.Name)).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                    Filename:=fullPath, _
                                    Quality:=xlQualityStandard, _
                                    IgnorePrintAreas:=False, _
                                    OpenAfterPublish:=False

    ' Clean up the temporary sheet
    Application.DisplayAlerts = False
    duplicateWs.Delete
    Application.DisplayAlerts = True

    ' Select the original sheet
    originalWs.Select

    MsgBox "Invoice exported successfully as a 2-page PDF!" & vbCrLf & _
           "File: " & cleanInvoiceNumber & ".pdf" & vbCrLf & _
           "Location: " & pdfPath, vbInformation, "PDF Export Complete"
    Exit Sub

ErrorHandler:
    ' Ensure cleanup happens even if there's an error
    If Not duplicateWs Is Nothing Then
        Application.DisplayAlerts = False
        duplicateWs.Delete
        Application.DisplayAlerts = True
    End If
    MsgBox "Error exporting PDF: " & Err.Description & vbCrLf & _
           "Please check if the folder path exists and you have write permissions.", vbCritical, "PDF Export Error"
End Sub

Public Sub PrintButton()
    ' Button function: Save as PDF and then send to default printer
    Dim ws As Worksheet
    Dim invoiceNumber As String
    Dim response As VbMsgBoxResult
    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets("GST_Tax_Invoice_for_interstate")

    ' Get invoice number
    invoiceNumber = Trim(ws.Range("C7").Value)

    If invoiceNumber = "" Then
        MsgBox "Please ensure invoice number is filled before printing.", vbExclamation, "Missing Invoice Number"
        Exit Sub
    End If

    ' First, save as PDF (call the PDF export function)
    Call PrintAsPDFButton

    ' Configure print settings
    With ws.PageSetup
        .PrintArea = "A1:O40"
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .CenterHorizontally = True
        .CenterVertically = False
        .PrintComments = xlPrintNoComments
        .PrintErrors = xlPrintErrorsDisplayed
    End With

    ' Confirm printing
    response = MsgBox("Send invoice " & invoiceNumber & " to printer?" & vbCrLf & _
                     "PDF has been saved to: /Users/narendrachowdary/BNC/gst invoices/", _
                     vbYesNo + vbQuestion, "Confirm Print")

    If response = vbYes Then
        ' Print the invoice
        ws.PrintOut Copies:=1, Preview:=False, ActivePrinter:=""

        MsgBox "Invoice " & invoiceNumber & " sent to printer successfully!" & vbCrLf & _
               "PDF copy saved to: /Users/narendrachowdary/BNC/gst invoices/", _
               vbInformation, "Print Complete"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error printing invoice: " & Err.Description, vbCritical, "Print Error"
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

        ' State dropdown for Receiver (Row 15, Column C15:F15)
        .Range("C15").Validation.Delete
        .Range("C15").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$J$2:$J$37"
        .Range("C15").Validation.IgnoreBlank = True
        .Range("C15").Validation.InCellDropdown = True
        .Range("C15").Validation.ShowError = False

        ' State dropdown for Consignee (Row 15, Column I15:K15)
        .Range("I15").Validation.Delete
        .Range("I15").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$J$2:$J$37"
        .Range("I15").Validation.IgnoreBlank = True
        .Range("I15").Validation.InCellDropdown = True
        .Range("I15").Validation.ShowError = False

        ' GSTIN dropdown for Receiver (Row 14, Column C14)
        .Range("C14").Validation.Delete
        .Range("C14").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$X$2:$X$2"
        .Range("C14").Validation.IgnoreBlank = True
        .Range("C14").Validation.InCellDropdown = True
        .Range("C14").Validation.ShowError = False

        ' GSTIN dropdown for Consignee (Row 14, Column I14)
        .Range("I14").Validation.Delete
        .Range("I14").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$X$2:$X$2"
        .Range("I14").Validation.IgnoreBlank = True
        .Range("I14").Validation.InCellDropdown = True
        .Range("I14").Validation.ShowError = False
    End With

    On Error GoTo 0
End Sub
' ===== MULTI-ITEM SUPPORT SYSTEM =====

Public Sub AddNewItemRow()
    Dim ws As Worksheet
    Dim lastItemRow As Long
    Dim newRowNum As Long
    Dim i As Integer

    On Error GoTo ErrorHandler

    Set ws = GetOrCreateWorksheet("GST_Tax_Invoice_for_interstate")

    If ws Is Nothing Then
        MsgBox "Invoice sheet not found!", vbExclamation
        Exit Sub
    End If

    ' Find the last item row (starts from row 18)
    lastItemRow = 18
    Do While ws.Cells(lastItemRow + 1, 1).Value <> "" Or lastItemRow < 22
        lastItemRow = lastItemRow + 1
        If lastItemRow > 22 Then Exit Do
    Loop

    ' Check if we can add more rows (limit to row 22)
    If lastItemRow >= 22 Then
        MsgBox "Maximum 5 items allowed in this invoice format!", vbExclamation
        Exit Sub
    End If

    newRowNum = lastItemRow + 1

    With ws
        ' Copy formatting from the previous row
        .Rows(lastItemRow).Copy
        .Rows(newRowNum).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False

        ' Clear the values but keep formatting
        .Range("A" & newRowNum & ":K" & newRowNum).ClearContents

        ' Set the Sr.No. for the new row
        .Cells(newRowNum, 1).Value = newRowNum - 17  ' Sr.No. starts from 1

        ' Setup formulas for the new row
        .Range("G" & newRowNum).Formula = "=IF(AND(D" & newRowNum & "<>"""",F" & newRowNum & "<>""""),D" & newRowNum & "*F" & newRowNum & ","""")"
        .Range("H" & newRowNum).Formula = "=IF(G" & newRowNum & "<>"""",G" & newRowNum & ","""")"
        .Range("I" & newRowNum).Value = "12"
        .Range("J" & newRowNum).Formula = "=IF(AND(H" & newRowNum & "<>"""",I" & newRowNum & "<>""""),H" & newRowNum & "*I" & newRowNum & "/100,"""")"
        .Range("K" & newRowNum).Formula = "=IF(AND(H" & newRowNum & "<>"""",J" & newRowNum & "<>""""),H" & newRowNum & "+J" & newRowNum & ","""")"

        ' Format the new row
        .Range("G" & newRowNum & ":K" & newRowNum).NumberFormat = "0.00"
        .Range("I" & newRowNum).NumberFormat = "0.00"
        .Rows(newRowNum).RowHeight = 32

        ' Update the summary calculations to include all rows
        Call UpdateMultiItemTaxCalculations(ws)
    End With

    MsgBox "New item row added successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error adding new item row: " & Err.Description, vbCritical
End Sub


' ===== BUTTON CREATION =====

Public Sub CreateInvoiceButtons(ws As Worksheet)
    ' Create professional buttons for invoice operations - Robust individual approach
    On Error GoTo ErrorHandler

    ' Remove any existing buttons first
    Call RemoveExistingButtons(ws)

    ' Add a small delay to ensure the worksheet is ready for button creation
    Application.Wait (Now + TimeValue("0:00:01"))
    
    ' Create buttons with cell-based positioning for better visibility
    Call CreateButtonAtCell(ws, "Q7", "Save Customer to Warehouse", "AddCustomerToWarehouseButton")
    Call CreateButtonAtCell(ws, "Q9", "Save Invoice Record", "SaveInvoiceButton")
    Call CreateButtonAtCell(ws, "Q11", "New Invoice", "NewInvoiceButton")
    Call CreateButtonAtCell(ws, "Q15", "Add New Item Row", "AddNewItemRowButton")
    Call CreateButtonAtCell(ws, "Q21", "Export as PDF", "PrintAsPDFButton")
    Call CreateButtonAtCell(ws, "Q23", "Print Invoice", "PrintButton")

    ' Add section headers
    Call CreateSectionHeaders(ws)

    ' Debug: Show button creation summary
    Debug.Print "Button creation completed. Total buttons in worksheet: " & ws.Buttons.Count

    Exit Sub

ErrorHandler:
    Debug.Print "Button creation error: " & Err.Description & " (Error #" & Err.Number & ")"
    MsgBox "Error creating invoice buttons: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Buttons created so far: " & ws.Buttons.Count & vbCrLf & _
           "This may be due to existing buttons or worksheet protection.", vbCritical, "Button Creation Error"
End Sub

Private Sub CreateButtonAtCell(ws As Worksheet, cellAddress As String, caption As String, onAction As String)
    ' Create a button positioned at a specific cell
    Dim btn As Button
    Dim targetCell As Range
    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btnWidth As Double
    Dim btnHeight As Double
    On Error Resume Next

    Set targetCell = ws.Range(cellAddress)

    ' Use cell position and size for button placement
    btnLeft = targetCell.Left
    btnTop = targetCell.Top
    btnWidth = 180  ' Fixed width
    btnHeight = 25  ' Fixed height

    Debug.Print "Creating button at " & cellAddress & " - Left: " & btnLeft & ", Top: " & btnTop

    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)

    If Err.Number = 0 And Not btn Is Nothing Then
        btn.Caption = caption
        btn.OnAction = onAction
        btn.Font.Name = "Segoe UI"
        btn.Font.Size = 9
        btn.Font.Bold = True

        ' Ensure button is on top
        btn.BringToFront

        Debug.Print "✅ Created button: " & caption & " at " & cellAddress
        
        ' Yield execution to allow Excel to process events
        DoEvents
    Else
        Debug.Print "❌ Failed to create button: " & caption & " - " & Err.Description
    End If

    Err.Clear
    On Error GoTo 0
End Sub

Private Sub CreateSectionHeaders(ws As Worksheet)
    ' Create section headers with individual error handling
    On Error Resume Next

    ' INVOICE OPERATIONS header
    ws.Range("Q6").Value = "INVOICE OPERATIONS"
    ws.Range("Q6").Font.Bold = True
    ws.Range("Q6").Font.Size = 11
    ws.Range("Q6").Font.Color = RGB(47, 80, 97)
    ws.Range("Q6").HorizontalAlignment = xlCenter

    ' ITEM MANAGEMENT header
    ws.Range("Q14").Value = "ITEM MANAGEMENT"
    ws.Range("Q14").Font.Bold = True
    ws.Range("Q14").Font.Size = 11
    ws.Range("Q14").Font.Color = RGB(47, 80, 97)
    ws.Range("Q14").HorizontalAlignment = xlCenter

    ' PRINT & EXPORT header
    ws.Range("Q20").Value = "PRINT & EXPORT"
    ws.Range("Q20").Font.Bold = True
    ws.Range("Q20").Font.Size = 11
    ws.Range("Q20").Font.Color = RGB(47, 80, 97)
    ws.Range("Q20").HorizontalAlignment = xlCenter

    ' Footer note
    ws.Range("Q25").Value = "Click buttons for quick operations"
    ws.Range("Q25").Font.Size = 8
    ws.Range("Q25").Font.Italic = True
    ws.Range("Q25").Font.Color = RGB(100, 100, 100)
    ws.Range("Q25").HorizontalAlignment = xlCenter

    On Error GoTo 0
End Sub

Private Sub RemoveExistingButtons(ws As Worksheet)
    ' Remove any existing buttons to prevent duplicates
    Dim btn As Button
    Dim i As Integer
    On Error Resume Next

    ' Clear all buttons in the worksheet (more reliable approach)
    Do While ws.Buttons.Count > 0
        ws.Buttons(1).Delete
    Loop

    On Error GoTo 0
End Sub