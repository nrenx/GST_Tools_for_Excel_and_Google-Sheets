Option Explicit
' ===============================================================================
' MODULE: Module_InvoiceStructure
' DESCRIPTION: Handles the creation, formatting, and layout of the invoice sheet,
'              as well as its core formulas and data population logic.
' ===============================================================================

' ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
' 📋 WORKSHEET CREATION FUNCTIONS
' ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

Public Sub CreateInvoiceSheet()
    Dim ws As Worksheet
    Dim i As Long
    Dim headers As Variant
    Dim basicHeaders As Variant
    Dim itemData As Variant
    Dim receiverData() As Variant
    Dim consigneeData() As Variant

    ' Suppress Excel alerts to prevent merge cells warning
    Application.DisplayAlerts = False

    ' --- Setup with comprehensive error handling ---
    On Error GoTo ErrorHandler

    ' Try to get the sheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("GST_Tax_Invoice_for_interstate")
    On Error GoTo 0

    ' If the sheet doesn't exist, create it. If it exists, clear it completely.
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "GST_Tax_Invoice_for_interstate"
    Else
        ' Complete cleanup of existing sheet
        On Error Resume Next
        ws.Cells.UnMerge
        ws.Cells.Clear
        On Error GoTo 0
    End If

    ' Activate the sheet safely
    On Error Resume Next
    ws.Activate
    On Error GoTo 0

    ' --- Main Formatting Block ---
    With ws
        ' Set column widths safely - EXPANDED LAYOUT A-P (16 columns)
        On Error Resume Next
        .Columns(1).ColumnWidth = 5    ' Column A - Sr.No.
        .Columns(2).ColumnWidth = 12   ' Column B - Description of Goods/Services
        .Columns(3).ColumnWidth = 12   ' Column C - HSN/SAC Code
        .Columns(4).ColumnWidth = 9    ' Column D - Quantity
        .Columns(5).ColumnWidth = 7    ' Column E - UOM
        .Columns(6).ColumnWidth = 10   ' Column F - Rate
        .Columns(7).ColumnWidth = 14   ' Column G - Amount
        .Columns(8).ColumnWidth = 10   ' Column H - Taxable Value
        .Columns(9).ColumnWidth = 6    ' Column I - IGST Rate
        .Columns(10).ColumnWidth = 10  ' Column J - IGST Amount
        .Columns(11).ColumnWidth = 6   ' Column K - CGST Rate
        .Columns(12).ColumnWidth = 10  ' Column L - CGST Amount
        .Columns(13).ColumnWidth = 6   ' Column M - SGST Rate
        .Columns(14).ColumnWidth = 10  ' Column N - SGST Amount
        .Columns(15).ColumnWidth = 12  ' Column O - Total Amount
        .Columns(16).ColumnWidth = 10  ' Column P - Sale Type
        On Error GoTo 0

        ' Set default font for all cells (before applying specific formatting)
        On Error Resume Next
        .Cells.Font.Name = "Segoe UI"
        .Cells.Font.Size = 11 ' Increased default font size
        .Cells.Font.Color = RGB(26, 26, 26)
        On Error GoTo 0

        ' Create header sections with premium professional styling - OPTIMIZED TO COLUMN O
        Call CreateHeaderRow(ws, 1, "A1:O1", "ORIGINAL", 12, True, RGB(47, 80, 97), RGB(255, 255, 255), 25)
        Call CreateHeaderRow(ws, 2, "A2:O2", "KAVERI TRADERS", 24, True, RGB(47, 80, 97), RGB(255, 255, 255), 37)
        Call CreateHeaderRow(ws, 3, "A3:O3", "191, Guduru, Pagadalapalli, Idulapalli, Tirupati, Andhra Pradesh - 524409", 11, True, RGB(245, 245, 245), RGB(26, 26, 26), 27)
        Call CreateHeaderRow(ws, 4, "A4:O4", "GSTIN: 37HERPB7733F1Z5", 14, True, RGB(245, 245, 245), RGB(26, 26, 26), 27)
        Call CreateHeaderRow(ws, 5, "A5:O5", "Email: kotidarisetty7777@gmail.com", 11, True, RGB(245, 245, 245), RGB(26, 26, 26), 25)

        ' Remove ONLY INTERNAL borders from rows 3 and 4 - PRESERVE OUTER BORDERS
        On Error Resume Next
        ' Remove internal borders from row 3 but preserve left and right outer borders
        .Range("A3:O3").Borders(xlInsideHorizontal).LineStyle = xlNone
        .Range("A3:O3").Borders(xlInsideVertical).LineStyle = xlNone
        .Range("A3:O3").Borders(xlEdgeTop).LineStyle = xlNone
        .Range("A3:O3").Borders(xlEdgeBottom).LineStyle = xlNone

        ' Remove internal borders from row 4 but preserve left and right outer borders
        .Range("A4:O4").Borders(xlInsideHorizontal).LineStyle = xlNone
        .Range("A4:O4").Borders(xlInsideVertical).LineStyle = xlNone
        .Range("A4:O4").Borders(xlEdgeTop).LineStyle = xlNone
        .Range("A4:O4").Borders(xlEdgeBottom).LineStyle = xlNone

        ' Also remove bottom border of row 2 (between row 2 and 3)
        .Range("A2:O2").Borders(xlEdgeBottom).LineStyle = xlNone
        On Error GoTo 0

        ' Row 6: TAX-INVOICE header - PROPERLY PROPORTIONED FOR OPTIMIZED LAYOUT
        Call CreateHeaderRow(ws, 6, "A6:J6", "TAX-INVOICE", 22, True, RGB(240, 240, 240), RGB(0, 0, 0), 28)
        Call CreateHeaderRow(ws, 6, "K6:O6", "Original for Recipient" & vbLf & "Duplicate for Supplier/Transporter" & vbLf & "Triplicate for Supplier", 9, True, RGB(250, 250, 250), RGB(0, 0, 0), 45)

        ' Enable text wrapping for the right section and ensure center alignment for TAX-INVOICE - PROPERLY PROPORTIONED
        On Error Resume Next
        .Range("A6:J6").HorizontalAlignment = xlCenter
        .Range("A6:J6").VerticalAlignment = xlCenter
        .Range("K6:O6").WrapText = True
        On Error GoTo 0

        ' --- Invoice Details with Merged Cells ---
        On Error Resume Next

        ' Row 7: Invoice No., Transport Mode, Challan No.
        .Range("A7:B7").Merge
        .Range("A7").Value = "Invoice No."
        .Range("A7").Font.Bold = True
        .Range("A7").HorizontalAlignment = xlLeft
        .Range("A7").Interior.Color = RGB(245, 245, 245)
        .Range("A7").Font.Color = RGB(26, 26, 26)
        .Range("C7").Value = ""
        .Range("C7").Font.Bold = True
        .Range("C7").Font.Color = RGB(220, 20, 60)  ' Red color for user input
        .Range("C7").HorizontalAlignment = xlCenter
        .Range("C7").VerticalAlignment = xlCenter

        .Range("D7:E7").Merge
        .Range("D7").Value = "Transport Mode"
        .Range("D7").Font.Bold = True
        .Range("D7").HorizontalAlignment = xlLeft
        .Range("D7").Interior.Color = RGB(245, 245, 245)
        .Range("D7").Font.Color = RGB(26, 26, 26)
        .Range("F7:G7").Merge
        .Range("F7").Value = "By Lorry"
        .Range("F7").HorizontalAlignment = xlLeft

        .Range("H7:I7").Merge
        .Range("H7").Value = "Challan No."
        .Range("H7").Font.Bold = True
        .Range("H7").HorizontalAlignment = xlLeft
        .Range("H7").Interior.Color = RGB(245, 245, 245)
        .Range("H7").Font.Color = RGB(26, 26, 26)
        .Range("J7:K7").Merge
        .Range("J7").Value = ""
        .Range("J7").HorizontalAlignment = xlLeft

        ' Row 8: Invoice Date, Vehicle Number, Transporter Name
        .Range("A8:B8").Merge
        .Range("A8").Value = "Invoice Date"
        .Range("A8").Font.Bold = True
        .Range("A8").HorizontalAlignment = xlLeft
        .Range("A8").Interior.Color = RGB(245, 245, 245)
        .Range("A8").Font.Color = RGB(26, 26, 26)
        .Range("C8").Value = ""
        .Range("C8").Font.Bold = True
        .Range("C8").HorizontalAlignment = xlLeft

        .Range("D8:E8").Merge
        .Range("D8").Value = "Vehicle Number"
        .Range("D8").Font.Bold = True
        .Range("D8").HorizontalAlignment = xlLeft
        .Range("D8").Interior.Color = RGB(245, 245, 245)
        .Range("D8").Font.Color = RGB(26, 26, 26)
        .Range("F8:G8").Merge
        .Range("F8").Value = ""
        .Range("F8").HorizontalAlignment = xlLeft

        .Range("H8:I8").Merge
        .Range("H8").Value = "Transporter Name"
        .Range("H8").Font.Bold = True
        .Range("H8").HorizontalAlignment = xlLeft
        .Range("H8").Interior.Color = RGB(245, 245, 245)
        .Range("H8").Font.Color = RGB(26, 26, 26)
        .Range("J8:K8").Merge
        .Range("J8").Value = "NARENDRA"
        .Range("J8").HorizontalAlignment = xlLeft

        ' Row 8: Additional fields for expanded layout (Columns L-P)
        .Range("L8:M8").Merge
        .Range("L8").Value = "Invoice Type"
        .Range("L8").Font.Bold = True
        .Range("L8").HorizontalAlignment = xlLeft
        .Range("L8").Interior.Color = RGB(245, 245, 245)
        .Range("L8").Font.Color = RGB(26, 26, 26)
        .Range("N8:O8").Merge
        .Range("N8").Value = "Tax Invoice"
        .Range("N8").HorizontalAlignment = xlCenter

        ' Row 9: State, Date of Supply, L.R Number
        .Range("A9:B9").Merge
        .Range("A9").Value = "State"
        .Range("A9").Font.Bold = True
        .Range("A9").HorizontalAlignment = xlLeft
        .Range("A9").Interior.Color = RGB(245, 245, 245)
        .Range("A9").Font.Color = RGB(26, 26, 26)
        .Range("C9").Value = "Andhra Pradesh"
        .Range("C9").HorizontalAlignment = xlLeft
        .Range("C9").Font.Size = 10

        .Range("D9:E9").Merge
        .Range("D9").Value = "Date of Supply"
        .Range("D9").Font.Bold = True
        .Range("D9").HorizontalAlignment = xlLeft
        .Range("D9").Interior.Color = RGB(245, 245, 245)
        .Range("D9").Font.Color = RGB(26, 26, 26)
        .Range("F9:G9").Merge
        .Range("F9").Value = ""
        .Range("F9").HorizontalAlignment = xlLeft

        .Range("H9:I9").Merge
        .Range("H9").Value = "L.R Number"
        .Range("H9").Font.Bold = True
        .Range("H9").HorizontalAlignment = xlLeft
        .Range("H9").Interior.Color = RGB(245, 245, 245)
        .Range("H9").Font.Color = RGB(26, 26, 26)
        .Range("J9:K9").Merge
        .Range("J9").Value = ""
        .Range("J9").HorizontalAlignment = xlLeft

        ' Row 9: Additional fields for expanded layout (Columns L-P)
        .Range("L9:M9").Merge
        .Range("L9").Value = "Reverse Charge"
        .Range("L9").Font.Bold = True
        .Range("L9").HorizontalAlignment = xlLeft
        .Range("L9").Interior.Color = RGB(245, 245, 245)
        .Range("L9").Font.Color = RGB(26, 26, 26)
        .Range("N9:O9").Merge
        .Range("N9").Value = "No"
        .Range("N9").HorizontalAlignment = xlCenter

        ' Row 10: State Code, Place of Supply, P.O Number
        .Range("A10:B10").Merge
        .Range("A10").Value = "State Code"
        .Range("A10").Font.Bold = True
        .Range("A10").HorizontalAlignment = xlLeft
        .Range("A10").Interior.Color = RGB(245, 245, 245)
        .Range("A10").Font.Color = RGB(26, 26, 26)
        .Range("C10").Value = "37"
        .Range("C10").HorizontalAlignment = xlLeft

        .Range("D10:E10").Merge
        .Range("D10").Value = "Place of Supply"
        .Range("D10").Font.Bold = True
        .Range("D10").HorizontalAlignment = xlLeft
        .Range("D10").Interior.Color = RGB(245, 245, 245)
        .Range("D10").Font.Color = RGB(26, 26, 26)
        .Range("F10:G10").Merge
        .Range("F10").Value = ""
        .Range("F10").HorizontalAlignment = xlLeft

        .Range("H10:I10").Merge
        .Range("H10").Value = "P.O Number"
        .Range("H10").Font.Bold = True
        .Range("H10").HorizontalAlignment = xlLeft
        .Range("H10").Interior.Color = RGB(245, 245, 245)
        .Range("H10").Font.Color = RGB(26, 26, 26)
        .Range("J10:K10").Merge
        .Range("J10").Value = ""
        .Range("J10").HorizontalAlignment = xlLeft

        ' Row 10: Additional fields for expanded layout (Columns L-P)
        .Range("L10:M10").Merge
        .Range("L10").Value = "E-Way Bill No."
        .Range("L10").Font.Bold = True
        .Range("L10").HorizontalAlignment = xlLeft
        .Range("L10").Interior.Color = RGB(245, 245, 245)
        .Range("L10").Font.Color = RGB(26, 26, 26)
        .Range("N10:O10").Merge
        .Range("N10").Value = ""
        .Range("N10").HorizontalAlignment = xlCenter

        ' NEW: Sale Type field (Row 7, Columns L-O)
        .Range("L7:M7").Merge
        .Range("L7").Value = "Sale Type"
        .Range("L7").Font.Bold = True
        .Range("L7").HorizontalAlignment = xlLeft
        .Range("L7").Interior.Color = RGB(245, 245, 245)
        .Range("L7").Font.Color = RGB(26, 26, 26)
        .Range("N7:O7").Merge
        .Range("N7").Value = "Interstate"  ' Default value
        .Range("N7").Font.Bold = True
        .Range("N7").Font.Color = RGB(220, 20, 60)  ' Red color for user input
        .Range("N7").HorizontalAlignment = xlCenter
        .Range("N7").VerticalAlignment = xlCenter

        ' Apply borders and formatting with professional color - OPTIMIZED TO COLUMN O
        .Range("A7:O10").Borders.LineStyle = xlContinuous
        .Range("A7:O10").Borders.Color = RGB(204, 204, 204)
        For i = 7 To 10
            .Rows(i).RowHeight = 28 ' Increased height
        Next i
        On Error GoTo 0

        ' --- Party Details (Professional Styling) - OPTIMIZED TO COLUMN O ---
        Call CreateHeaderRow(ws, 11, "A11:H11", "Details of Receiver (Billed to)", 11, True, RGB(245, 245, 245), RGB(26, 26, 26), 26)
        Call CreateHeaderRow(ws, 11, "I11:O11", "Details of Consignee (Shipped to)", 11, True, RGB(245, 245, 245), RGB(26, 26, 26), 26)

        ' Set center alignment for row 11 content (both horizontal and vertical)
        On Error Resume Next
        .Range("A11:H11").HorizontalAlignment = xlCenter
        .Range("A11:H11").VerticalAlignment = xlCenter
        .Range("I11:O11").HorizontalAlignment = xlCenter
        .Range("I11:O11").VerticalAlignment = xlCenter
        On Error GoTo 0

        ' --- Party Details with Merged Cells ---
        On Error Resume Next

        ' Row 12: Name fields - EXPANDED TO COLUMN P
        .Range("A12:B12").Merge
        .Range("A12").Value = "Name:"
        .Range("A12").Font.Bold = True
        .Range("A12").HorizontalAlignment = xlLeft
        .Range("A12").Interior.Color = RGB(245, 245, 245)
        .Range("A12").Font.Color = RGB(26, 26, 26)
        .Range("C12:H12").Merge
        .Range("C12").Value = ""
        .Range("C12").HorizontalAlignment = xlLeft

        .Range("I12:J12").Merge
        .Range("I12").Value = "Name:"
        .Range("I12").Font.Bold = True
        .Range("I12").HorizontalAlignment = xlLeft
        .Range("I12").Interior.Color = RGB(245, 245, 245)
        .Range("I12").Font.Color = RGB(26, 26, 26)
        .Range("K12:O12").Merge
        .Range("K12").Value = ""
        .Range("K12").HorizontalAlignment = xlLeft

        ' Row 13: Address fields - OPTIMIZED TO COLUMN O
        .Range("A13:B13").Merge
        .Range("A13").Value = "Address:"
        .Range("A13").Font.Bold = True
        .Range("A13").HorizontalAlignment = xlLeft
        .Range("A13").Interior.Color = RGB(245, 245, 245)
        .Range("A13").Font.Color = RGB(26, 26, 26)
        .Range("C13:H13").Merge
        .Range("C13").Value = ""
        .Range("C13").HorizontalAlignment = xlLeft

        .Range("I13:J13").Merge
        .Range("I13").Value = "Address:"
        .Range("I13").Font.Bold = True
        .Range("I13").HorizontalAlignment = xlLeft
        .Range("I13").Interior.Color = RGB(245, 245, 245)
        .Range("I13").Font.Color = RGB(26, 26, 26)
        .Range("K13:O13").Merge
        .Range("K13").Value = ""
        .Range("K13").HorizontalAlignment = xlLeft

        ' Row 14: GSTIN fields - OPTIMIZED TO COLUMN O
        .Range("A14:B14").Merge
        .Range("A14").Value = "GSTIN:"
        .Range("A14").Font.Bold = True
        .Range("A14").HorizontalAlignment = xlLeft
        .Range("A14").Interior.Color = RGB(245, 245, 245)
        .Range("A14").Font.Color = RGB(26, 26, 26)
        .Range("C14:H14").Merge
        .Range("C14").Value = ""
        .Range("C14").HorizontalAlignment = xlLeft

        .Range("I14:J14").Merge
        .Range("I14").Value = "GSTIN:"
        .Range("I14").Font.Bold = True
        .Range("I14").HorizontalAlignment = xlLeft
        .Range("I14").Interior.Color = RGB(245, 245, 245)
        .Range("I14").Font.Color = RGB(26, 26, 26)
        .Range("K14:O14").Merge
        .Range("K14").Value = ""
        .Range("K14").HorizontalAlignment = xlLeft

        ' Row 15: State fields - OPTIMIZED TO COLUMN O
        .Range("A15:B15").Merge
        .Range("A15").Value = "State:"
        .Range("A15").Font.Bold = True
        .Range("A15").HorizontalAlignment = xlLeft
        .Range("A15").Interior.Color = RGB(245, 245, 245)
        .Range("A15").Font.Color = RGB(26, 26, 26)
        .Range("C15:H15").Merge
        .Range("C15").Value = ""
        .Range("C15").HorizontalAlignment = xlLeft

        .Range("I15:J15").Merge
        .Range("I15").Value = "State:"
        .Range("I15").Font.Bold = True
        .Range("I15").HorizontalAlignment = xlLeft
        .Range("I15").Interior.Color = RGB(245, 245, 245)
        .Range("I15").Font.Color = RGB(26, 26, 26)
        .Range("K15:O15").Merge
        .Range("K15").Value = ""
        .Range("K15").HorizontalAlignment = xlLeft

        ' Row 16: State Code fields - OPTIMIZED TO COLUMN O
        .Range("A16:B16").Merge
        .Range("A16").Value = "State Code:"
        .Range("A16").Font.Bold = True
        .Range("A16").HorizontalAlignment = xlLeft
        .Range("A16").Interior.Color = RGB(245, 245, 245)
        .Range("A16").Font.Color = RGB(26, 26, 26)
        .Range("C16:H16").Merge
        .Range("C16").Formula = "=VLOOKUP(C15, warehouse!J2:K37, 2, FALSE)"
        .Range("C16").HorizontalAlignment = xlLeft

        .Range("I16:J16").Merge
        .Range("I16").Value = "State Code:"
        .Range("I16").Font.Bold = True
        .Range("I16").HorizontalAlignment = xlLeft
        .Range("I16").Interior.Color = RGB(245, 245, 245)
        .Range("I16").Font.Color = RGB(26, 26, 26)
        .Range("K16:O16").Merge
        .Range("K16").Formula = "=VLOOKUP(K15, warehouse!J2:K37, 2, FALSE)"
        .Range("K16").HorizontalAlignment = xlLeft

        ' Apply borders and formatting for rows 12-16 with professional color - OPTIMIZED TO COLUMN O
        .Range("A12:O16").Borders.LineStyle = xlContinuous
        .Range("A12:O16").Borders.Color = RGB(204, 204, 204)
        For i = 12 To 16
            .Rows(i).RowHeight = 28 ' Increased height
        Next i
        On Error GoTo 0

        ' --- Item Table (Simplified) ---
        On Error Resume Next

        ' TWO-ROW HEADER STRUCTURE WITH PROPER TAX COLUMN MERGING
        On Error Resume Next

    ' STEP 1: Create individual cell headers first (before merging)
    ' Row 17: Set header texts (will merge non-tax columns vertically)
    .Cells(17, 1).Value = "Sr.No."
    .Cells(17, 2).Value = "Description of Goods/Services"
    .Cells(17, 3).Value = "HSN/SAC Code"
    .Cells(17, 4).Value = "Quantity"
    .Cells(17, 5).Value = "UOM"
    .Cells(17, 6).Value = "Rate"
    .Cells(17, 7).Value = "Amount"
    .Cells(17, 8).Value = "Taxable Value"
    .Cells(17, 15).Value = "Total Amount"

    ' Row 18: Set only tax column sub-headers; non-tax will be merged vertically
    .Cells(18, 1).Value = ""
    .Cells(18, 2).Value = ""
    .Cells(18, 3).Value = ""
    .Cells(18, 4).Value = ""
    .Cells(18, 5).Value = ""
    .Cells(18, 6).Value = ""
    .Cells(18, 7).Value = ""
    .Cells(18, 8).Value = ""
    .Cells(18, 9).Value = "Rate (%)"
    .Cells(18, 10).Value = "Amount (Rs.)"
    .Cells(18, 11).Value = "Rate (%)"
    .Cells(18, 12).Value = "Amount (Rs.)"
    .Cells(18, 13).Value = "Rate (%)"
    .Cells(18, 14).Value = "Amount (Rs.)"
    .Cells(18, 15).Value = ""

        ' STEP 2: Apply formatting to all header cells
        .Range("A17:O18").Font.Bold = True
        .Range("A17:O17").Font.Size = 10
        .Range("A18:O18").Font.Size = 9
        .Range("A17:O17").Interior.Color = RGB(245, 245, 245)
        .Range("A18:O18").Interior.Color = RGB(250, 250, 250)
        .Range("A17:O18").Font.Color = RGB(26, 26, 26)
        .Range("A17:O18").WrapText = True
        .Range("A17:O18").HorizontalAlignment = xlCenter
        .Range("A17:O18").VerticalAlignment = xlCenter
        .Range("A17:O18").Borders.LineStyle = xlContinuous
        .Range("A17:O18").Borders.Color = RGB(204, 204, 204)

        ' STEP 3: Merge non-tax columns vertically across Rows 17-18 for consistent two-row header
        .Range("A17:A18").Merge
        .Range("B17:B18").Merge
        .Range("C17:C18").Merge
        .Range("D17:D18").Merge
        .Range("E17:E18").Merge
        .Range("F17:F18").Merge
        .Range("G17:G18").Merge
        .Range("H17:H18").Merge
        .Range("O17:O18").Merge

        ' Ensure merged non-tax columns use the light gray background and centered alignment
        With .Range("A17:A18,B17:B18,C17:C18,D17:D18,E17:E18,F17:F18,G17:G18,H17:H18,O17:O18")
            .Interior.Color = RGB(245, 245, 245)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ' STEP 4: Create merged cells for tax columns in Row 17 (unchanged)
        ' STEP 3: Create merged cells for tax columns in Row 17
        ' CGST: Merge columns I,J (9,10)
        .Range("I17:J17").Merge
        .Range("I17").Value = "CGST"
        .Range("I17").HorizontalAlignment = xlCenter
        .Range("I17").VerticalAlignment = xlCenter

        ' SGST: Merge columns K,L (11,12)
        .Range("K17:L17").Merge
        .Range("K17").Value = "SGST"
        .Range("K17").HorizontalAlignment = xlCenter
        .Range("K17").VerticalAlignment = xlCenter

        ' IGST: Merge columns M,N (13,14)
        .Range("M17:N17").Merge
        .Range("M17").Value = "IGST"
        .Range("M17").HorizontalAlignment = xlCenter
        .Range("M17").VerticalAlignment = xlCenter

    ' Set optimal row heights for two-row header
        .Rows(17).RowHeight = 30
        .Rows(18).RowHeight = 30

        On Error GoTo 0

        ' Item data - ENHANCED STRUCTURE (ROW 19, COLUMNS A-O) - UPDATED FOR TWO-ROW HEADER
        itemData = Array("1", "Casuarina Wood", "", "", "", "", "", "", "", "", "", "", "", "", "")
        For i = 0 To UBound(itemData)
            With .Cells(19, i + 1)
                .Value = itemData(i)
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(204, 204, 204)
                .Font.Size = 10
                .Interior.Color = RGB(250, 250, 250)
                If i = 0 Or i = 2 Or i = 3 Or i = 4 Then  ' Sr.No, HSN, Qty, UOM
                    .HorizontalAlignment = xlCenter
                ElseIf i = 1 Then  ' Description
                    .HorizontalAlignment = xlLeft
                ElseIf i >= 5 Then  ' Rate through Total Amount (columns F-O)
                    .HorizontalAlignment = xlRight
                    .Font.Bold = True
                End If
            End With
        Next i
        .Rows(19).RowHeight = 35

        ' Setup automatic tax calculation formulas
        Call SetupTaxCalculationFormulas(ws)

        ' Empty rows with alternating colors - ENHANCED STRUCTURE (6 PRODUCT ROWS) - UPDATED FOR TWO-ROW HEADER
        For i = 20 To 24  ' Rows 20-24 (item rows 2-6), row 25 is totals
            .Range("A" & i & ":O" & i).Borders.LineStyle = xlContinuous
            .Range("A" & i & ":O" & i).Borders.Color = RGB(204, 204, 204)
            If i Mod 2 = 0 Then
                .Range("A" & i & ":O" & i).Interior.Color = RGB(250, 250, 250)
            Else
                .Range("A" & i & ":O" & i).Interior.Color = RGB(255, 255, 255)
            End If
            .Rows(i).RowHeight = 30 ' Increased height
        Next i
        On Error GoTo 0

        ' --- Row 25 Total Quantity Section - ENHANCED STRUCTURE (6 PRODUCT ROWS) - UPDATED FOR TWO-ROW HEADER ---
        On Error Resume Next

        ' Apply borders to entire row first
        .Range("A25:O25").Borders.LineStyle = xlContinuous
        .Range("A25:O25").Borders.Color = RGB(204, 204, 204)
        .Range("A25:O25").Interior.Color = RGB(234, 234, 234)
        .Rows(25).RowHeight = 30

        ' Merge A25:C25 for "Total Quantity" label
        .Range("A25:C25").Merge
        .Range("A25").Value = "Total Quantity"
        .Range("A25").Font.Bold = True
        .Range("A25").HorizontalAlignment = xlCenter
        .Range("A25").VerticalAlignment = xlCenter
        .Range("A25").Font.Color = RGB(26, 26, 26)

        ' Cell D25 for quantity value
        .Range("D25").Value = ""
        .Range("D25").Font.Bold = True
        .Range("D25").HorizontalAlignment = xlCenter

        ' Merge E25:F25 for "Sub Total" label
        .Range("E25:F25").Merge
        .Range("E25").Value = "Sub Total:"
        .Range("E25").Font.Bold = True
        .Range("E25").HorizontalAlignment = xlRight
        .Range("E25").Font.Color = RGB(26, 26, 26)

        ' Individual cells for amounts (G, H for Amount and Taxable Value)
        .Range("G25").Value = ""
        .Range("G25").Font.Bold = True
        .Range("G25").HorizontalAlignment = xlRight

        .Range("H25").Value = ""
        .Range("H25").Font.Bold = True
        .Range("H25").HorizontalAlignment = xlRight

        ' Tax amount cells - ENHANCED STRUCTURE (I-N for all tax types)
        .Range("I25").Value = ""  ' IGST Rate
        .Range("I25").Font.Bold = True
        .Range("I25").HorizontalAlignment = xlRight

        .Range("J25").Value = ""  ' IGST Amount
        .Range("J25").Font.Bold = True
        .Range("J25").HorizontalAlignment = xlRight

        .Range("K25").Value = ""  ' CGST Rate
        .Range("K25").Font.Bold = True
        .Range("K25").HorizontalAlignment = xlRight

        .Range("L25").Value = ""  ' CGST Amount
        .Range("L25").Font.Bold = True
        .Range("L25").HorizontalAlignment = xlRight

        .Range("M25").Value = ""  ' SGST Rate
        .Range("M25").Font.Bold = True
        .Range("M25").HorizontalAlignment = xlRight

        .Range("N25").Value = ""  ' SGST Amount
        .Range("N25").Font.Bold = True
        .Range("N25").HorizontalAlignment = xlRight

        ' Cell O25 for total amount
        .Range("O25").Value = ""
        .Range("O25").Font.Bold = True
        .Range("O25").HorizontalAlignment = xlRight

        On Error GoTo 0

        ' --- Row 26-28 Total Invoice Amount in Words Section - ENHANCED STRUCTURE - UPDATED FOR TWO-ROW HEADER ---
        On Error Resume Next

        ' Row 26: Header for "Total Invoice Amount in Words"
        .Range("A26:J26").Merge
        .Range("A26").Value = "Total Invoice Amount in Words"
        .Range("A26").Font.Bold = True
        .Range("A26").Font.Size = 13 ' Increased font size
        .Range("A26").HorizontalAlignment = xlCenter
        .Range("A26").Interior.Color = RGB(255, 255, 0)
        .Range("A26:J26").Borders.LineStyle = xlContinuous
        .Rows(26).RowHeight = 25

        ' Rows 27-28: Amount in words content (merged across 2 rows)
        .Range("A27:J28").Merge
        .Range("A27").Value = ""
        .Range("A27").Font.Bold = True
        .Range("A27").Font.Size = 15 ' Increased font size
        .Range("A27").HorizontalAlignment = xlCenter
        .Range("A27").VerticalAlignment = xlCenter
        .Range("A27").Interior.Color = RGB(255, 255, 230)
        .Range("A27").Borders.LineStyle = xlContinuous
        .Range("A27").WrapText = True
        .Rows(27).RowHeight = 25 ' Increased height
        .Rows(28).RowHeight = 25 ' Increased height

        ' Tax summary on the right (columns K-O, rows 26-32) - ENHANCED STRUCTURE - UPDATED FOR TWO-ROW HEADER
        ' FIXED POSITIONING: Tax summary starts at row 26 to avoid overlap

        ' Row 26: Total Before Tax
        .Range("K26:N26").Merge
        .Range("K26").Value = "Total Amount Before Tax:"
        .Range("K26").Font.Bold = True
        .Range("K26").Font.Size = 11 ' Increased font size
        .Range("K26").HorizontalAlignment = xlLeft
        .Range("K26").Interior.Color = RGB(245, 245, 245)
        .Range("K26").Font.Color = RGB(26, 26, 26)

        .Range("O26").Value = ""
        .Range("O26").Font.Bold = True
        .Range("O26").HorizontalAlignment = xlRight
        .Range("O26").Interior.Color = RGB(216, 222, 233)

        ' Row 27: CGST
        .Range("K27:N27").Merge
        .Range("K27").Value = "CGST :"
        .Range("K27").Font.Bold = True
        .Range("K27").Font.Size = 11 ' Increased font size
        .Range("K27").HorizontalAlignment = xlLeft
        .Range("K27").Interior.Color = RGB(245, 245, 245)
        .Range("K27").Font.Color = RGB(26, 26, 26)

        .Range("O27").Value = ""
        .Range("O27").Font.Bold = True
        .Range("O27").HorizontalAlignment = xlRight
        .Range("O27").Interior.Color = RGB(216, 222, 233)

        ' Row 28: SGST
        .Range("K28:N28").Merge
        .Range("K28").Value = "SGST :"
        .Range("K28").Font.Bold = True
        .Range("K28").Font.Size = 11 ' Increased font size
        .Range("K28").HorizontalAlignment = xlLeft
        .Range("K28").Interior.Color = RGB(245, 245, 245)
        .Range("K28").Font.Color = RGB(26, 26, 26)

        .Range("O28").Value = ""
        .Range("O28").Font.Bold = True
        .Range("O28").HorizontalAlignment = xlRight
        .Range("O28").Interior.Color = RGB(216, 222, 233)

        ' Row 29: IGST (highlighted)
        .Range("K29:N29").Merge
        .Range("K29").Value = "IGST :"
        .Range("K29").Font.Bold = True
        .Range("K29").Font.Size = 11 ' Increased font size
        .Range("K29").HorizontalAlignment = xlLeft
        .Range("K29").Interior.Color = RGB(255, 255, 200)
        .Range("K29").Font.Color = RGB(26, 26, 26)

        .Range("O29").Value = ""
        .Range("O29").Font.Bold = True
        .Range("O29").HorizontalAlignment = xlRight
        .Range("O29").Interior.Color = RGB(255, 255, 200)

        ' Row 30: CESS
        .Range("K30:N30").Merge
        .Range("K30").Value = "CESS :"
        .Range("K30").Font.Bold = True
        .Range("K30").Font.Size = 11 ' Increased font size
        .Range("K30").HorizontalAlignment = xlLeft
        .Range("K30").Interior.Color = RGB(245, 245, 245)
        .Range("K30").Font.Color = RGB(26, 26, 26)

        .Range("O30").Value = ""
        .Range("O30").Font.Bold = True
        .Range("O30").HorizontalAlignment = xlRight
        .Range("O30").Interior.Color = RGB(216, 222, 233)

        ' Row 31: Total Tax (highlighted) - ENHANCED STRUCTURE - UPDATED FOR TWO-ROW HEADER
        .Range("K31:N31").Merge
        With .Range("K31")
            .Value = "Total Tax:"
            .Font.Bold = True
            .Font.Size = 11 ' Increased font size
            .Interior.Color = RGB(240, 240, 240)
            .Font.Color = RGB(26, 26, 26)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With

        .Range("O31").Value = ""
        .Range("O31").Font.Bold = True
        .Range("O31").HorizontalAlignment = xlRight
        .Range("O31").Interior.Color = RGB(240, 240, 240)

        ' Rows 32-33: Total Amount After Tax - ENHANCED TWO-ROW PROMINENCE MATCHING REFERENCE LAYOUT
        ' Merge K32:N33 for the label
        .Range("K32:N33").Merge
        .Range("K32").Value = "Total Amount After Tax:"
        .Range("K32").Font.Bold = True
        .Range("K32").Font.Size = 12 ' Larger font size for prominence
        .Range("K32").HorizontalAlignment = xlCenter
        .Range("K32").VerticalAlignment = xlCenter
        .Range("K32").Interior.Color = RGB(255, 215, 0)  ' Gold background for prominence
        .Range("K32").Font.Color = RGB(0, 0, 0)  ' Black text for contrast

        ' Merge O32:O33 for the calculated value
        .Range("O32:O33").Merge
        .Range("O32").Value = ""  ' Will be populated by formula
        .Range("O32").Font.Bold = True
        .Range("O32").Font.Size = 12 ' Larger font size for prominence
        .Range("O32").HorizontalAlignment = xlCenter
        .Range("O32").VerticalAlignment = xlCenter
        .Range("O32").Interior.Color = RGB(255, 215, 0)  ' Gold background for prominence
        .Range("O32").Font.Color = RGB(0, 0, 0)  ' Black text for contrast

        ' Set row heights for two-row prominence
        .Rows(32).RowHeight = 30
        .Rows(33).RowHeight = 30

        ' Set row heights for tax summary section - UPDATED FOR TWO-ROW HEADER
        .Rows(26).RowHeight = 20
        .Rows(27).RowHeight = 20
        .Rows(28).RowHeight = 20
        .Rows(29).RowHeight = 20
        .Rows(30).RowHeight = 20
        .Rows(31).RowHeight = 20
        .Rows(32).RowHeight = 20

        On Error GoTo 0

        ' Setup automatic tax calculation formulas for summary section
        Call UpdateMultiItemTaxCalculations(ws)

        ' --- Terms and Conditions Section - MOVED TO CORRECT POSITION ---
        ' Row 29: Header for "Terms and Conditions"
        .Range("A29:J29").Merge
        .Range("A29").Value = "Terms and Conditions"
        .Range("A29").Font.Bold = True
        .Range("A29").Font.Size = 13
        .Range("A29").HorizontalAlignment = xlCenter
        .Range("A29").Interior.Color = RGB(255, 255, 0)  ' Yellow background like reference
        .Range("A29").Borders.LineStyle = xlContinuous
        .Rows(29).RowHeight = 25

        ' Rows 30-33: Terms and conditions content (merged across 4 rows for better spacing)
        .Range("A30:J33").Merge
        .Range("A30").Value = "1. This is an electronically generated invoice." & vbLf & _
                             "2. All disputes are subject to GUDUR jurisdiction only." & vbLf & _
                             "3. If the Consignee makes any Inter State Sale, he has to pay GST himself." & vbLf & _
                             "4. Goods once sold cannot be taken back or exchanged." & vbLf & _
                             "5. Payment terms as per agreement between buyer and seller."
        .Range("A30").Font.Size = 11
        .Range("A30").HorizontalAlignment = xlLeft
        .Range("A30").VerticalAlignment = xlTop
        .Range("A30").Interior.Color = RGB(255, 255, 245)  ' Light yellow background
        .Range("A30").Borders.LineStyle = xlContinuous
        .Range("A30").WrapText = True
        For i = 30 To 33
            .Rows(i).RowHeight = 25
        Next i

        ' Amount in words conversion integrated into signature section

        On Error GoTo 0

        ' --- Signature Section with Merged Cells - ENHANCED STRUCTURE - UPDATED FOR EXTRA TERMS ROW ---
        On Error Resume Next

        ' Row 34: Signature headers with merged cells - ENHANCED STRUCTURE - UPDATED FOR EXTRA TERMS ROW
        .Range("A34:E34").Merge
        .Range("A34").Value = "Transporter"
        .Range("A34").Font.Bold = True
        .Range("A34").HorizontalAlignment = xlCenter
        .Range("A34").Interior.Color = RGB(220, 220, 220)

        .Range("F34:J34").Merge
        .Range("F34").Value = "Receiver"
        .Range("F34").Font.Bold = True
        .Range("F34").HorizontalAlignment = xlCenter
        .Range("F34").Interior.Color = RGB(220, 220, 220)

        .Range("K34:O34").Merge
        .Range("K34").Value = "Certified that the particulars given above are true and correct"
        .Range("K34").Font.Bold = True
        .Range("K34").Font.Size = 10 ' Increased font size
        .Range("K34").HorizontalAlignment = xlCenter
        .Range("K34").VerticalAlignment = xlCenter
        .Range("K34").WrapText = True
        .Range("K34").Interior.Color = RGB(220, 220, 220)

        ' Rows 35-36: Mobile Number Section (merged across 2 rows) - UPDATED FOR EXTRA TERMS ROW
        .Range("A35:E36").Merge
        .Range("A35").Value = "Mobile No: ___________________"
        .Range("A35").Font.Size = 10 ' Increased font size
        .Range("A35").HorizontalAlignment = xlCenter
        .Range("A35").VerticalAlignment = xlCenter
        .Range("A35").Interior.Color = RGB(250, 250, 250)

        .Range("F35:J36").Merge
        .Range("F35").Value = "Mobile No: ___________________"
        .Range("F35").Font.Size = 10 ' Increased font size
        .Range("F35").HorizontalAlignment = xlCenter
        .Range("F35").VerticalAlignment = xlCenter
        .Range("F35").Interior.Color = RGB(250, 250, 250)

        .Range("K35:O36").Merge
        .Range("K35").Value = "Mobile No: ___________________"
        .Range("K35").Font.Size = 10 ' Increased font size
        .Range("K35").HorizontalAlignment = xlCenter
        .Range("K35").VerticalAlignment = xlCenter
        .Range("K35").Interior.Color = RGB(250, 250, 250)

        ' Rows 37-39: Signature Space Section (merged across 3 rows) - UPDATED FOR EXTRA TERMS ROW
        .Range("A37:E39").Merge
        .Range("A37").Value = ""
        .Range("A37").Interior.Color = RGB(250, 250, 250)

        .Range("F37:J39").Merge
        .Range("F37").Value = ""
        .Range("F37").Interior.Color = RGB(250, 250, 250)

        .Range("K37:O39").Merge
        .Range("K37").Value = ""
        .Range("K37").Interior.Color = RGB(250, 250, 250)

        ' Row 40: Signature Labels - UPDATED FOR EXTRA TERMS ROW
        .Range("A40:E40").Merge
        .Range("A40").Value = "Transporter's Signature"
        .Range("A40").Font.Bold = True
        .Range("A40").Font.Size = 10 ' Increased font size
        .Range("A40").HorizontalAlignment = xlCenter
        .Range("A40").Interior.Color = RGB(211, 211, 211)

        .Range("F40:J40").Merge
        .Range("F40").Value = "Receiver's Signature"
        .Range("F40").Font.Bold = True
        .Range("F40").Font.Size = 10 ' Increased font size
        .Range("F40").HorizontalAlignment = xlCenter
        .Range("F40").Interior.Color = RGB(211, 211, 211)

        .Range("K40:O40").Merge
        .Range("K40").Value = "Authorized Signatory"
        .Range("K40").Font.Bold = True
        .Range("K40").Font.Size = 10 ' Increased font size
        .Range("K40").HorizontalAlignment = xlCenter
        .Range("K40").Interior.Color = RGB(211, 211, 211)

        ' Set specific row height for row 34 to accommodate wrapped text
        .Rows(34).RowHeight = 35

        ' Set standard row height for remaining signature rows
        For i = 35 To 40
            .Rows(i).RowHeight = 25 ' Increased height
        Next i
        .Rows(39).RowHeight = 31
        On Error GoTo 0
    End With

    ' --- Final Formatting ---
    With ws
        On Error Resume Next
        ' Font settings moved to beginning of code to avoid overriding header fonts
        On Error GoTo 0

        ' Apply professional page setup - OPTIMIZED FOR ENHANCED LAYOUT AND SCALING
        On Error Resume Next
        With .PageSetup
            .PrintArea = "A1:O40"  ' Set print area to include all enhanced content
            .Orientation = xlPortrait
            .PaperSize = xlPaperA4
            .Zoom = False ' Let Excel handle scaling
            .FitToPagesWide = 1
            .FitToPagesTall = 1 ' Fit to one page vertically
            .LeftMargin = Application.InchesToPoints(0.15)  ' Reduced margins for more content space
            .RightMargin = Application.InchesToPoints(0.15)
            .TopMargin = Application.InchesToPoints(0.15)
            .BottomMargin = Application.InchesToPoints(0.15)
            .HeaderMargin = Application.InchesToPoints(0.1)
            .FooterMargin = Application.InchesToPoints(0.1)
            .CenterHorizontally = True
            .CenterVertically = True  ' Enable vertical centering for better appearance
        End With
        On Error GoTo 0

        ' COMPREHENSIVE BORDER FIX - Apply consistent borders to entire invoice - UPDATED FOR EXTRA TERMS ROW
        On Error Resume Next

        ' First, clear all existing borders to prevent conflicts
        .Range("A1:O40").Borders.LineStyle = xlNone

        ' Apply consistent internal borders to entire invoice area
        .Range("A1:O40").Borders.LineStyle = xlContinuous
        .Range("A1:O40").Borders.Weight = xlThin
        .Range("A1:O40").Borders.Color = RGB(0, 0, 0)  ' Pure black for PDF

        ' Apply outer border around entire invoice
        With .Range("A1:O40")
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeRight).Color = RGB(0, 0, 0)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeTop).Color = RGB(0, 0, 0)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
        End With

        ' FINAL STEP: Remove ONLY INTERNAL borders from rows 3 and 4 - PRESERVE OUTER BORDERS
        ' Remove internal horizontal and vertical borders but keep left and right outer borders
        .Range("A3:O3").Borders(xlInsideHorizontal).LineStyle = xlNone
        .Range("A3:O3").Borders(xlInsideVertical).LineStyle = xlNone
        .Range("A3:O3").Borders(xlEdgeTop).LineStyle = xlNone
        .Range("A3:O3").Borders(xlEdgeBottom).LineStyle = xlNone

        .Range("A4:O4").Borders(xlInsideHorizontal).LineStyle = xlNone
        .Range("A4:O4").Borders(xlInsideVertical).LineStyle = xlNone
        .Range("A4:O4").Borders(xlEdgeTop).LineStyle = xlNone
        .Range("A4:O4").Borders(xlEdgeBottom).LineStyle = xlNone

        ' Also remove bottom border of row 2 for seamless header appearance
        .Range("A2:O2").Borders(xlEdgeBottom).LineStyle = xlNone

        On Error GoTo 0
    End With

    ' Create professional buttons for invoice operations
    Call CreateInvoiceButtons(ws)

    ' Auto-populate invoice number and dates
    Call AutoPopulateInvoiceFields(ws)

    ' Set up worksheet change events for auto-population
    Call SetupWorksheetChangeEvents(ws)

    ' Set up worksheet change events for state code extraction
    Call SetupStateCodeChangeEvents(ws)

    ' Auto-fill consignee from receiver data
    Call AutoFillConsigneeFromReceiver(ws)

    ' Setup dynamic tax display based on default sale type
    Call SetupDynamicTaxDisplay(ws)

    ' Setup data validation dropdowns (including Sale Type dropdown)
    Call SetupDataValidation(ws)

    ' Restore Excel alerts
    Application.DisplayAlerts = True

    MsgBox "GST TAX-INVOICE created successfully with expanded layout and dynamic tax functionality!", vbInformation
    Exit Sub

ErrorHandler:
    ' Restore Excel alerts even in case of error
    Application.DisplayAlerts = True
    MsgBox "An error occurred in CreateInvoiceSheet." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "Error"
    On Error GoTo 0
End Sub

Private Sub CreateHeaderRow(ws As Worksheet, rowNum As Integer, rangeAddr As String, text As String, fontSize As Integer, isBold As Boolean, backColor As Long, fontColor As Long, rowHeight As Integer)
    On Error Resume Next

    ' Set the text in the first cell
    ws.Range(rangeAddr).Cells(1, 1).Value = CleanText(text)

    ' Apply formatting to the entire range
    With ws.Range(rangeAddr)
        .Font.Bold = isBold
        .Font.Size = fontSize
        .Font.Color = fontColor
        .Interior.Color = backColor
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium

        ' Try to merge - if it fails, continue anyway
        .Merge
        If Err.Number <> 0 Then Err.Clear
    End With

    ' Set row height
    ws.Rows(rowNum).RowHeight = rowHeight

    On Error GoTo 0
End Sub

Private Sub AutoPopulateInvoiceFields(ws As Worksheet)
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
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        ' Allow manual editing - no validation restrictions
    End With

    With ws.Range("G9")
        .Value = Format(Date, "dd/mm/yyyy")
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        ' Allow manual editing - no validation restrictions
    End With

    ' Set fixed State Code (Row 10, Column C) for Andhra Pradesh
    With ws.Range("C10")
        .Value = "37"
        .Font.Bold = True
        .Interior.Color = RGB(245, 245, 245)  ' Light grey background
        .Font.Color = RGB(26, 26, 26)  ' Dark text
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

Private Sub SetupWorksheetChangeEvents(ws As Worksheet)
    ' Set up change monitoring for customer dropdown auto-population and Sale Type handling
    ' Since we're in a module, we'll ensure the worksheet change events are enabled
    On Error Resume Next

    ' Enable automatic calculation to ensure formulas update properly
    Application.Calculation = xlCalculationAutomatic
    
    ' Note: The actual worksheet change events for Sale Type are handled by 
    ' the HandleSaleTypeChange function in Module_InvoiceEvents
    ' This function can be called manually when the Sale Type changes

    On Error GoTo 0
End Sub

Private Sub SetupStateCodeChangeEvents(ws As Worksheet)
    ' Simple state code setup - no automatic extraction needed
    ' State code dropdowns will show simple numeric codes only
    On Error GoTo 0
End Sub

Private Sub AutoFillConsigneeFromReceiver(ws As Worksheet)
    ' Automatically copy all receiver data to consignee fields - UPDATED FOR EXPANDED LAYOUT
    On Error GoTo ErrorHandler

    With ws
        ' Copy Name from Receiver (C12:H12) to Consignee (K12:O12)
        .Range("K12").Value = .Range("C12").Value

        ' Copy Address from Receiver (C13:H13) to Consignee (K13:O13)
        .Range("K13").Value = .Range("C13").Value

        ' Copy GSTIN from Receiver (C14:H14) to Consignee (K14:O14)
        .Range("K14").Value = .Range("C14").Value

        ' Copy State from Receiver (C15:H15) to Consignee (K15:O15)
        .Range("K15").Value = .Range("C15").Value

        ' State code for consignee is now handled by a VLOOKUP formula in cell K16.
        ' This line is no longer needed as copying the state name to K15 will trigger the formula.

        ' Format consignee fields for manual editing (use default black font)
        .Range("K12:O12").Font.Color = RGB(26, 26, 26)  ' Standard black font
        .Range("K13:O13").Font.Color = RGB(26, 26, 26)  ' Standard black font
        .Range("K14:O14").Font.Color = RGB(26, 26, 26)  ' Standard black font
        .Range("K15:O15").Font.Color = RGB(26, 26, 26)  ' Standard black font
        .Range("K16:O16").Font.Color = RGB(26, 26, 26)  ' Standard black font
    End With

    Exit Sub

ErrorHandler:
    ' If auto-fill fails, continue silently
    On Error GoTo 0
End Sub

' 🧮 TAX CALCULATION FUNCTIONS

Public Sub SetupTaxCalculationFormulas(ws As Worksheet)
    ' Set up formulas for automatic tax calculations in the item table with enhanced structure - UPDATED FOR TWO-ROW HEADER
    On Error Resume Next

    With ws
        ' For row 19 (first item row), set up formulas - ENHANCED STRUCTURE A-O - UPDATED FOR TWO-ROW HEADER
        ' Column G (Amount) = Quantity * Rate
        .Range("G19").Formula = "=IF(AND(D19<>"""",F19<>""""),D19*F19,"""")"

        ' Column H (Taxable Value) = Amount (same as Amount for simplicity)
        .Range("H19").Formula = "=IF(G19<>"""",G19,"""")"

        ' Column I (CGST Rate) - VLOOKUP formula to get tax rate from HSN data (half of total rate for intrastate)
        .Range("I19").Formula = "=IF(N7=""Intrastate"",VLOOKUP(C19, warehouse!A:E, 5, FALSE)/2,"""")"

        ' Column J (CGST Amount) = Taxable Value * CGST Rate / 100
        .Range("J19").Formula = "=IF(AND(H19<>"""",I19<>""""),H19*I19/100,"""")"

        ' Column K (SGST Rate) - VLOOKUP formula to get tax rate from HSN data (half of total rate for intrastate)
        .Range("K19").Formula = "=IF(N7=""Intrastate"",VLOOKUP(C19, warehouse!A:E, 5, FALSE)/2,"""")"

        ' Column L (SGST Amount) = Taxable Value * SGST Rate / 100
        .Range("L19").Formula = "=IF(AND(H19<>"""",K19<>""""),H19*K19/100,"""")"

        ' Column M (IGST Rate) - VLOOKUP formula to get tax rate from HSN data (only for interstate)
        .Range("M19").Formula = "=IF(N7=""Interstate"",VLOOKUP(C19, warehouse!A:E, 5, FALSE),"""")"

        ' Column N (IGST Amount) = Taxable Value * IGST Rate / 100 (only for interstate)
        .Range("N19").Formula = "=IF(AND(H19<>"""",M19<>""""),H19*M19/100,"""")"

        ' Column O (Total Amount) = Taxable Value + Tax Amounts (IGST for interstate, CGST+SGST for intrastate)
        .Range("O19").Formula = "=IF(N7=""Interstate"",H19+N19,IF(N7=""Intrastate"",H19+J19+L19,H19))"

        ' Format the formula cells - ENHANCED STRUCTURE
        .Range("G19:O19").NumberFormat = "0.00"
        .Range("I19,K19,M19").NumberFormat = "0.00"
    End With

    On Error GoTo 0
End Sub

Public Sub UpdateMultiItemTaxCalculations(ws As Worksheet)
    ' Update tax calculations to sum all item rows with enhanced structure - UPDATED FOR TWO-ROW HEADER
    On Error Resume Next

    With ws
        ' Row 25: Total Quantity - ENHANCED STRUCTURE - UPDATED FOR TWO-ROW HEADER
        .Range("D25").Formula = "=SUM(D19:D24)"
        .Range("D25").NumberFormat = "#,##0.00"

        ' Row 25: Sub Total calculations
        .Range("G25").Formula = "=SUM(G19:G24)"  ' Amount column
        .Range("H25").Formula = "=SUM(H19:H24)"  ' Taxable Value column
        .Range("G25:H25").NumberFormat = "#,##0.00"

        ' Row 25: Tax amounts - ENHANCED STRUCTURE - UPDATED FOR CORRECT COLUMN ORDER (CGST, SGST, IGST)
        .Range("I25").Formula = "=SUM(I19:I24)"  ' CGST Rate (average)
        .Range("J25").Formula = "=SUM(J19:J24)"  ' CGST Amount
        .Range("K25").Formula = "=SUM(K19:K24)"  ' SGST Rate (average)
        .Range("L25").Formula = "=SUM(L19:L24)"  ' SGST Amount
        .Range("M25").Formula = "=SUM(M19:M24)"  ' IGST Rate (average)
        .Range("N25").Formula = "=SUM(N19:N24)"  ' IGST Amount
        .Range("O25").Formula = "=SUM(O19:O24)"  ' Total Amount
        .Range("I25:O25").NumberFormat = "#,##0.00"

        ' Tax summary section (right side) - ENHANCED STRUCTURE - UPDATED FOR CORRECT COLUMN ORDER
        ' Row 26: Total Amount Before Tax
        .Range("O26").Formula = "=SUM(H19:H24)"

        ' Row 27: CGST
        .Range("O27").Formula = "=SUM(J19:J24)"

        ' Row 28: SGST
        .Range("O28").Formula = "=SUM(L19:L24)"

        ' Row 29: IGST
        .Range("O29").Formula = "=SUM(N19:N24)"

        ' Row 30: CESS (0 by default)
        .Range("O30").Value = 0

        ' Row 31: Total Tax
        .Range("O31").Formula = "=O27+O28+O29+O30"

        ' Row 32: Total Amount After Tax
        .Range("O32").Formula = "=O26+O31"

        ' Format all calculation cells
        .Range("O26:O32").NumberFormat = "#,##0.00"

        ' Update Amount in Words (A27 merged cell) - ENHANCED STRUCTURE - UPDATED FOR TWO-ROW HEADER
        .Range("A27").Formula = "=NumberToWords(O32)"
    End With

    On Error GoTo 0
End Sub

Public Sub SetupDynamicTaxDisplay(ws As Worksheet)
    ' Set up dynamic tax field display based on sale type
    On Error Resume Next

    With ws
        ' Set up conditional formatting for "Not Applicable" display
        ' This will be handled through worksheet change events

        ' Initialize with default Interstate setup
        Call UpdateTaxFieldsDisplay(ws, "Interstate")
    End With

    On Error GoTo 0
End Sub

Public Sub UpdateTaxFieldsDisplay(ws As Worksheet, saleType As String)
    ' Update tax fields display based on sale type selection - FIXED COLUMN MAPPING
    Dim i As Long
    On Error Resume Next

    With ws
        If saleType = "Interstate" Then
            ' INTERSTATE: Only IGST applies, CGST and SGST are not applicable
            
            ' Clear all tax fields first for all 6 product rows (19-24)
            .Range("I19:N24").ClearContents
            
            ' Restore proper headers for active IGST columns (M,N)
            .Range("M17").Value = "IGST Rate (%)"
            .Range("M17").Font.Color = RGB(26, 26, 26)  ' Black color
            .Range("M17").Font.Bold = True
            .Range("M17").HorizontalAlignment = xlCenter
            
            .Range("N17").Value = "IGST Amount (Rs.)"
            .Range("N17").Font.Color = RGB(26, 26, 26)  ' Black color
            .Range("N17").Font.Bold = True
            .Range("N17").HorizontalAlignment = xlCenter
            
            ' Set "Not Apply" messages in red for CGST and SGST headers
            .Range("I17").Value = "CGST Not Apply"
            .Range("I17").Font.Color = RGB(220, 20, 60)  ' Red color
            .Range("I17").Font.Bold = True
            .Range("I17").HorizontalAlignment = xlCenter
            
            .Range("J17").Value = "CGST Not Apply"
            .Range("J17").Font.Color = RGB(220, 20, 60)  ' Red color
            .Range("J17").Font.Bold = True
            .Range("J17").HorizontalAlignment = xlCenter
            
            .Range("K17").Value = "SGST Not Apply"
            .Range("K17").Font.Color = RGB(220, 20, 60)  ' Red color
            .Range("K17").Font.Bold = True
            .Range("K17").HorizontalAlignment = xlCenter
            
            .Range("L17").Value = "SGST Not Apply"
            .Range("L17").Font.Color = RGB(220, 20, 60)  ' Red color
            .Range("L17").Font.Bold = True
            .Range("L17").HorizontalAlignment = xlCenter
            
            ' Clear content completely from CGST columns (I19-I24, J19-J24)
            .Range("I19:I24").ClearContents
            .Range("J19:J24").ClearContents
            
            ' Clear content completely from SGST columns (K19-K24, L19-L24)
            .Range("K19:K24").ClearContents
            .Range("L19:L24").ClearContents
            
            ' Set up active IGST formulas (M,N columns)
            For i = 19 To 24
                .Range("M" & i).Formula = "=IF(AND(N7=""Interstate"",C" & i & "<>""""),IFERROR(VLOOKUP(C" & i & ", warehouse!A:E, 5, FALSE),""""),"""")"
                .Range("N" & i).Formula = "=IF(AND(N7=""Interstate"",H" & i & "<>"""",M" & i & "<>""""),H" & i & "*M" & i & "/100,"""")"
            Next i

        ElseIf saleType = "Intrastate" Then
            ' INTRASTATE: Only CGST and SGST apply, IGST is not applicable
            
            ' Clear all tax fields first for all 6 product rows (19-24)
            .Range("I19:N24").ClearContents
            
            ' Restore proper headers for active CGST columns (I,J)
            .Range("I17").Value = "CGST Rate (%)"
            .Range("I17").Font.Color = RGB(26, 26, 26)  ' Black color
            .Range("I17").Font.Bold = True
            .Range("I17").HorizontalAlignment = xlCenter
            
            .Range("J17").Value = "CGST Amount (Rs.)"
            .Range("J17").Font.Color = RGB(26, 26, 26)  ' Black color
            .Range("J17").Font.Bold = True
            .Range("J17").HorizontalAlignment = xlCenter
            
            ' Restore proper headers for active SGST columns (K,L)
            .Range("K17").Value = "SGST Rate (%)"
            .Range("K17").Font.Color = RGB(26, 26, 26)  ' Black color
            .Range("K17").Font.Bold = True
            .Range("K17").HorizontalAlignment = xlCenter
            
            .Range("L17").Value = "SGST Amount (Rs.)"
            .Range("L17").Font.Color = RGB(26, 26, 26)  ' Black color
            .Range("L17").Font.Bold = True
            .Range("L17").HorizontalAlignment = xlCenter
            
            ' Set "Not Apply" messages in red for IGST headers
            .Range("M17").Value = "IGST Not Apply"
            .Range("M17").Font.Color = RGB(220, 20, 60)  ' Red color
            .Range("M17").Font.Bold = True
            .Range("M17").HorizontalAlignment = xlCenter
            
            .Range("N17").Value = "IGST Not Apply"
            .Range("N17").Font.Color = RGB(220, 20, 60)  ' Red color
            .Range("N17").Font.Bold = True
            .Range("N17").HorizontalAlignment = xlCenter
            
            ' Clear content completely from IGST columns (M19-M24, N19-N24)
            .Range("M19:M24").ClearContents
            .Range("N19:N24").ClearContents
            
            ' Set up active CGST formulas (I,J columns) - half of total GST rate
            For i = 19 To 24
                .Range("I" & i).Formula = "=IF(AND(N7=""Intrastate"",C" & i & "<>""""),IFERROR(VLOOKUP(C" & i & ", warehouse!A:E, 5, FALSE)/2,""""),"""")"
                .Range("J" & i).Formula = "=IF(AND(N7=""Intrastate"",H" & i & "<>"""",I" & i & "<>""""),H" & i & "*I" & i & "/100,"""")"
            Next i
            
            ' Set up active SGST formulas (K,L columns) - half of total GST rate
            For i = 19 To 24
                .Range("K" & i).Formula = "=IF(AND(N7=""Intrastate"",C" & i & "<>""""),IFERROR(VLOOKUP(C" & i & ", warehouse!A:E, 5, FALSE)/2,""""),"""")"
                .Range("L" & i).Formula = "=IF(AND(N7=""Intrastate"",H" & i & "<>"""",K" & i & "<>""""),H" & i & "*K" & i & "/100,"""")"
            Next i
        End If
        
        ' Force recalculation
        .Calculate
    End With

    On Error GoTo 0
End Sub

Public Sub CleanEmptyProductRows(ws As Worksheet)
    ' Clean up empty product rows to remove any #N/A values or unwanted content
    Dim i As Long
    On Error Resume Next

    With ws
        For i = 19 To 23  ' Product rows
            ' If no product description, clear the entire row
            If Trim(.Range("B" & i).Value) = "" And Trim(.Range("C" & i).Value) = "" Then
                .Range("A" & i & ":O" & i).ClearContents
                ' Set default formatting for empty rows
                .Range("A" & i & ":O" & i).Font.Color = RGB(26, 26, 26)
                .Range("A" & i & ":O" & i).Font.Bold = False
                .Range("A" & i & ":O" & i).Font.Size = 10
                .Range("A" & i & ":O" & i).HorizontalAlignment = xlLeft
                .Range("A" & i & ":O" & i).VerticalAlignment = xlCenter
            End If
        Next i
    End With

    On Error GoTo 0
End Sub