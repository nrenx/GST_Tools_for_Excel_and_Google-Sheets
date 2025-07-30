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
        ' Set column widths safely
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
        .Columns(11).ColumnWidth = 12  ' Column K - Total Amount
        On Error GoTo 0

        ' Set default font for all cells (before applying specific formatting)
        On Error Resume Next
        .Cells.Font.Name = "Segoe UI"
        .Cells.Font.Size = 11 ' Increased default font size
        .Cells.Font.Color = RGB(26, 26, 26)
        On Error GoTo 0

        ' Create header sections with premium professional styling
        Call CreateHeaderRow(ws, 1, "A1:K1", "ORIGINAL", 12, True, RGB(47, 80, 97), RGB(255, 255, 255), 25)
        Call CreateHeaderRow(ws, 2, "A2:K2", "KAVERI TRADERS", 24, True, RGB(47, 80, 97), RGB(255, 255, 255), 37)
        Call CreateHeaderRow(ws, 3, "A3:K3", "191, Guduru, Pagadalapalli, Idulapalli, Tirupati, Andhra Pradesh - 524409", 11, True, RGB(245, 245, 245), RGB(26, 26, 26), 27)
        Call CreateHeaderRow(ws, 4, "A4:K4", "GSTIN: 37HERPB7733F1Z5", 14, True, RGB(245, 245, 245), RGB(26, 26, 26), 27)
        Call CreateHeaderRow(ws, 5, "A5:K5", "Email: kotidarisetty7777@gmail.com", 11, True, RGB(245, 245, 245), RGB(26, 26, 26), 25)

        ' Remove horizontal borders between header rows for print/PDF appearance
        On Error Resume Next
        ' Remove bottom border of row 2 (between row 2 and 3)
        .Range("A2:K2").Borders(xlEdgeBottom).LineStyle = xlNone
        ' Remove bottom border of row 3 (between row 3 and 4)
        .Range("A3:K3").Borders(xlEdgeBottom).LineStyle = xlNone
        ' Remove bottom border of row 4 (between row 4 and 5)
        .Range("A4:K4").Borders(xlEdgeBottom).LineStyle = xlNone
        On Error GoTo 0

        ' Row 6: TAX-INVOICE header
        Call CreateHeaderRow(ws, 6, "A6:G6", "TAX-INVOICE", 22, True, RGB(240, 240, 240), RGB(0, 0, 0), 28)
        Call CreateHeaderRow(ws, 6, "H6:K6", "Original for Recipient" & vbLf & "Duplicate for Supplier/Transporter" & vbLf & "Triplicate for Supplier", 9, True, RGB(250, 250, 250), RGB(0, 0, 0), 45)

        ' Enable text wrapping for the right section and ensure center alignment for TAX-INVOICE
        On Error Resume Next
        .Range("A6:G6").HorizontalAlignment = xlCenter
        .Range("A6:G6").VerticalAlignment = xlCenter
        .Range("H6:K6").WrapText = True
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

        ' Apply borders and formatting with professional color
        .Range("A7:K10").Borders.LineStyle = xlContinuous
        .Range("A7:K10").Borders.Color = RGB(204, 204, 204)
        For i = 7 To 10
            .Rows(i).RowHeight = 28 ' Increased height
        Next i
        On Error GoTo 0

        ' --- Party Details (Professional Styling) ---
        Call CreateHeaderRow(ws, 11, "A11:F11", "Details of Receiver (Billed to)", 11, True, RGB(245, 245, 245), RGB(26, 26, 26), 26)
        Call CreateHeaderRow(ws, 11, "G11:K11", "Details of Consignee (Shipped to)", 11, True, RGB(245, 245, 245), RGB(26, 26, 26), 26)

        ' Set center alignment for row 11 content (both horizontal and vertical)
        On Error Resume Next
        .Range("A11:F11").HorizontalAlignment = xlCenter
        .Range("A11:F11").VerticalAlignment = xlCenter
        .Range("G11:K11").HorizontalAlignment = xlCenter
        .Range("G11:K11").VerticalAlignment = xlCenter
        On Error GoTo 0

        ' --- Party Details with Merged Cells ---
        On Error Resume Next

        ' Row 12: Name fields
        .Range("A12:B12").Merge
        .Range("A12").Value = "Name:"
        .Range("A12").Font.Bold = True
        .Range("A12").HorizontalAlignment = xlLeft
        .Range("A12").Interior.Color = RGB(245, 245, 245)
        .Range("A12").Font.Color = RGB(26, 26, 26)
        .Range("C12:F12").Merge
        .Range("C12").Value = ""
        .Range("C12").HorizontalAlignment = xlLeft

        .Range("G12:H12").Merge
        .Range("G12").Value = "Name:"
        .Range("G12").Font.Bold = True
        .Range("G12").HorizontalAlignment = xlLeft
        .Range("G12").Interior.Color = RGB(245, 245, 245)
        .Range("G12").Font.Color = RGB(26, 26, 26)
        .Range("I12:K12").Merge
        .Range("I12").Value = ""
        .Range("I12").HorizontalAlignment = xlLeft

        ' Row 13: Address fields
        .Range("A13:B13").Merge
        .Range("A13").Value = "Address:"
        .Range("A13").Font.Bold = True
        .Range("A13").HorizontalAlignment = xlLeft
        .Range("A13").Interior.Color = RGB(245, 245, 245)
        .Range("A13").Font.Color = RGB(26, 26, 26)
        .Range("C13:F13").Merge
        .Range("C13").Value = ""
        .Range("C13").HorizontalAlignment = xlLeft

        .Range("G13:H13").Merge
        .Range("G13").Value = "Address:"
        .Range("G13").Font.Bold = True
        .Range("G13").HorizontalAlignment = xlLeft
        .Range("G13").Interior.Color = RGB(245, 245, 245)
        .Range("G13").Font.Color = RGB(26, 26, 26)
        .Range("I13:K13").Merge
        .Range("I13").Value = ""
        .Range("I13").HorizontalAlignment = xlLeft

        ' Row 14: GSTIN fields
        .Range("A14:B14").Merge
        .Range("A14").Value = "GSTIN:"
        .Range("A14").Font.Bold = True
        .Range("A14").HorizontalAlignment = xlLeft
        .Range("A14").Interior.Color = RGB(245, 245, 245)
        .Range("A14").Font.Color = RGB(26, 26, 26)
        .Range("C14:F14").Merge
        .Range("C14").Value = ""
        .Range("C14").HorizontalAlignment = xlLeft

        .Range("G14:H14").Merge
        .Range("G14").Value = "GSTIN:"
        .Range("G14").Font.Bold = True
        .Range("G14").HorizontalAlignment = xlLeft
        .Range("G14").Interior.Color = RGB(245, 245, 245)
        .Range("G14").Font.Color = RGB(26, 26, 26)
        .Range("I14:K14").Merge
        .Range("I14").Value = ""
        .Range("I14").HorizontalAlignment = xlLeft

        ' Row 15: State fields
        .Range("A15:B15").Merge
        .Range("A15").Value = "State:"
        .Range("A15").Font.Bold = True
        .Range("A15").HorizontalAlignment = xlLeft
        .Range("A15").Interior.Color = RGB(245, 245, 245)
        .Range("A15").Font.Color = RGB(26, 26, 26)
        .Range("C15:F15").Merge
        .Range("C15").Value = ""
        .Range("C15").HorizontalAlignment = xlLeft

        .Range("G15:H15").Merge
        .Range("G15").Value = "State:"
        .Range("G15").Font.Bold = True
        .Range("G15").HorizontalAlignment = xlLeft
        .Range("G15").Interior.Color = RGB(245, 245, 245)
        .Range("G15").Font.Color = RGB(26, 26, 26)
        .Range("I15:K15").Merge
        .Range("I15").Value = ""
        .Range("I15").HorizontalAlignment = xlLeft

        ' Row 16: State Code fields
        .Range("A16:B16").Merge
        .Range("A16").Value = "State Code:"
        .Range("A16").Font.Bold = True
        .Range("A16").HorizontalAlignment = xlLeft
        .Range("A16").Interior.Color = RGB(245, 245, 245)
        .Range("A16").Font.Color = RGB(26, 26, 26)
        .Range("C16:F16").Merge
        .Range("C16").Formula = "=VLOOKUP(C15, warehouse!J2:K37, 2, FALSE)"
        .Range("C16").HorizontalAlignment = xlLeft

        .Range("G16:H16").Merge
        .Range("G16").Value = "State Code:"
        .Range("G16").Font.Bold = True
        .Range("G16").HorizontalAlignment = xlLeft
        .Range("G16").Interior.Color = RGB(245, 245, 245)
        .Range("G16").Font.Color = RGB(26, 26, 26)
        .Range("I16:K16").Merge
        .Range("I16").Formula = "=VLOOKUP(I15, warehouse!J2:K37, 2, FALSE)"
        .Range("I16").HorizontalAlignment = xlLeft

        ' Apply borders and formatting for rows 12-16 with professional color
        .Range("A12:K16").Borders.LineStyle = xlContinuous
        .Range("A12:K16").Borders.Color = RGB(204, 204, 204)
        For i = 12 To 16
            .Rows(i).RowHeight = 28 ' Increased height
        Next i
        On Error GoTo 0

        ' --- Item Table (Simplified) ---
        On Error Resume Next

        ' Headers
        headers = Array("Sr.No.", "Description of Goods/Services", "HSN/SAC Code", "Quantity", "UOM", "Rate (Rs.)", "Amount (Rs.)", "Taxable Value (Rs.)", "IGST Rate (%)", "IGST Amount (Rs.)", "Total Amount (Rs.)")
        For i = 0 To UBound(headers)
            With .Cells(17, i + 1)
                .Value = headers(i)
                .Font.Bold = True
                .Font.Size = 10 ' Increased font size
                .Interior.Color = RGB(245, 245, 245)
                .Font.Color = RGB(26, 26, 26)
                .WrapText = True
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(204, 204, 204)
            End With
        Next i
        .Rows(17).RowHeight = 58

        ' Item data
        itemData = Array("1", "Casuarina Wood", "", "", "", "", "", "", "", "", "")
        For i = 0 To UBound(itemData)
            With .Cells(18, i + 1)
                .Value = itemData(i)
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(204, 204, 204)
                .Font.Size = 10 ' Increased font size
                .Interior.Color = RGB(250, 250, 250)
                If i = 0 Or i = 2 Or i = 3 Or i = 4 Then
                    .HorizontalAlignment = xlCenter
                ElseIf i = 1 Then
                    .HorizontalAlignment = xlLeft
                ElseIf i >= 5 Then
                    .HorizontalAlignment = xlRight
                    .Font.Bold = True
                End If
            End With
        Next i
        .Rows(18).RowHeight = 35 ' Increased height

        ' Setup automatic tax calculation formulas
        Call SetupTaxCalculationFormulas(ws)

        ' Empty rows with alternating colors
        For i = 19 To 22
            .Range("A" & i & ":K" & i).Borders.LineStyle = xlContinuous
            .Range("A" & i & ":K" & i).Borders.Color = RGB(204, 204, 204)
            If i Mod 2 = 0 Then
                .Range("A" & i & ":K" & i).Interior.Color = RGB(250, 250, 250)
            Else
                .Range("A" & i & ":K" & i).Interior.Color = RGB(255, 255, 255)
            End If
            .Rows(i).RowHeight = 30 ' Increased height
        Next i
        On Error GoTo 0

        ' --- Row 22 Total Quantity Section (Restored) ---
        On Error Resume Next

        ' Merge A22:C22 for "Total Quantity" label
        .Range("A22:C22").Merge
        .Range("A22").Value = "Total Quantity"
        .Range("A22").Font.Bold = True
        .Range("A22").HorizontalAlignment = xlCenter
        .Range("A22").VerticalAlignment = xlBottom
        .Range("A22").Interior.Color = RGB(234, 234, 234)
        .Range("A22").Font.Color = RGB(26, 26, 26)

        ' Cell D22 for quantity value
        .Range("D22").Value = ""
        .Range("D22").Font.Bold = True
        .Range("D22").HorizontalAlignment = xlCenter
        .Range("D22").Interior.Color = RGB(234, 234, 234)

        ' Merge E22:F22 for "Sub Total" label
        .Range("E22:F22").Merge
        .Range("E22").Value = "Sub Total:"
        .Range("E22").Font.Bold = True
        .Range("E22").HorizontalAlignment = xlRight
        .Range("E22").Interior.Color = RGB(234, 234, 234)
        .Range("E22").Font.Color = RGB(26, 26, 26)

        ' Individual cells for amounts
        .Range("G22").Value = ""
        .Range("G22").Font.Bold = True
        .Range("G22").HorizontalAlignment = xlRight
        .Range("G22").Interior.Color = RGB(234, 234, 234)

        .Range("H22").Value = ""
        .Range("H22").Font.Bold = True
        .Range("H22").HorizontalAlignment = xlRight
        .Range("H22").Interior.Color = RGB(234, 234, 234)

        ' Merge I22:J22 for IGST amount
        .Range("I22:J22").Merge
        .Range("I22").Value = ""
        .Range("I22").Font.Bold = True
        .Range("I22").HorizontalAlignment = xlRight

        ' Cell K22 for total amount
        .Range("K22").Value = ""
        .Range("K22").Font.Bold = True
        .Range("K22").HorizontalAlignment = xlRight

        ' Apply formatting to the entire row with professional styling
        .Range("A22:K22").Interior.Color = RGB(234, 234, 234)
        .Range("A22:K22").Borders.LineStyle = xlContinuous
        .Range("A22:K22").Borders.Color = RGB(204, 204, 204)
        .Rows(22).RowHeight = 26
        On Error GoTo 0

        ' --- Row 23-25 Total Invoice Amount in Words Section (Restored) ---
        On Error Resume Next

        ' Row 23: Header for "Total Invoice Amount in Words"
        .Range("A23:G23").Merge
        .Range("A23").Value = "Total Invoice Amount in Words"
        .Range("A23").Font.Bold = True
        .Range("A23").Font.Size = 13 ' Increased font size
        .Range("A23").HorizontalAlignment = xlCenter
        .Range("A23").Interior.Color = RGB(255, 255, 0)
        .Range("A23:G23").Borders.LineStyle = xlContinuous
        .Rows(23).RowHeight = 25

        ' Rows 24-25: Amount in words content (merged across 2 rows)
        .Range("A24:G25").Merge
        .Range("A24").Value = ""
        .Range("A24").Font.Bold = True
        .Range("A24").Font.Size = 15 ' Increased font size
        .Range("A24").HorizontalAlignment = xlCenter
        .Range("A24").VerticalAlignment = xlCenter
        .Range("A24").Interior.Color = RGB(255, 255, 230)
        .Range("A24").Borders.LineStyle = xlContinuous
        .Range("A24").WrapText = True
        .Rows(24).RowHeight = 25 ' Increased height
        .Rows(25).RowHeight = 25 ' Increased height

        ' Tax summary on the right (columns H-K, rows 23-25) with merged cells

        ' Row 23: Total Before Tax
        .Range("H23:J23").Merge
        .Range("H23").Value = "Total Amount Before Tax:"
        .Range("H23").Font.Bold = True
        .Range("H23").Font.Size = 11 ' Increased font size
        .Range("H23").HorizontalAlignment = xlLeft
        .Range("H23").Interior.Color = RGB(245, 245, 245)
        .Range("H23").Font.Color = RGB(26, 26, 26)

        .Range("K23").Value = ""
        .Range("K23").Font.Bold = True
        .Range("K23").HorizontalAlignment = xlRight
        .Range("K23").Interior.Color = RGB(216, 222, 233)

        ' Row 24: CGST @ 0%
        .Range("H24:J24").Merge
        .Range("H24").Value = "CGST :"
        .Range("H24").Font.Bold = True
        .Range("H24").Font.Size = 11 ' Increased font size
        .Range("H24").HorizontalAlignment = xlLeft
        .Range("H24").Interior.Color = RGB(245, 245, 245)
        .Range("H24").Font.Color = RGB(26, 26, 26)

        .Range("K24").Value = ""
        .Range("K24").Font.Bold = True
        .Range("K24").HorizontalAlignment = xlRight
        .Range("K24").Interior.Color = RGB(216, 222, 233)

        ' Row 25: SGST @ 0%
        .Range("H25:J25").Merge
        .Range("H25").Value = "SGST :"
        .Range("H25").Font.Bold = True
        .Range("H25").Font.Size = 11 ' Increased font size
        .Range("H25").HorizontalAlignment = xlLeft
        .Range("H25").Interior.Color = RGB(245, 245, 245)
        .Range("H25").Font.Color = RGB(26, 26, 26)

        .Range("K25").Value = ""
        .Range("K25").Font.Bold = True
        .Range("K25").HorizontalAlignment = xlRight
        .Range("K25").Interior.Color = RGB(216, 222, 233)

        ' Apply borders to amount in words and tax summary sections
        .Range("H23:K25").Borders.LineStyle = xlContinuous
        .Range("H23:K25").Borders.Color = RGB(204, 204, 204)
        On Error GoTo 0

        ' Row 26: Header for "Terms and Conditions"
        .Range("A26:G26").Merge
        .Range("A26").Value = "Terms and Conditions"
        .Range("A26").Font.Bold = True
        .Range("A26").Font.Size = 13 ' Increased font size
        .Range("A26").HorizontalAlignment = xlCenter
        .Range("A26").Interior.Color = RGB(255, 255, 0)
        .Range("A26:G26").Borders.LineStyle = xlContinuous
        .Rows(26).RowHeight = 25

        ' Rows 27-30: Terms and conditions content (merged across 4 rows)
        .Range("A27:G30").Merge
        .Range("A27").Value = "1. This is an electronically generated invoice." & vbLf & _
                             "2. All disputes are subject to GUDUR jurisdiction only." & vbLf & _
                             "3. If the Consignee makes any Inter State Sales, he has to pay GST himself." & vbLf & _
                              "4. Goods once sold cannot be taken back or exchanged." & vbLf & _
                              "5. Payment terms: As per agreement between buyer and seller."
        .Range("A27").Font.Size = 11 ' Increased font size
        .Range("A27").HorizontalAlignment = xlLeft
        .Range("A27").VerticalAlignment = xlTop
        .Range("A27").Interior.Color = RGB(255, 255, 245)
        .Range("A27").Borders.LineStyle = xlContinuous
        .Range("A27").WrapText = True
        For i = 27 To 30
            .Rows(i).RowHeight = 25 ' Increased height
        Next i

        ' Tax summary on the right (columns H-K, rows 26-30) with merged cells

        ' Row 26: IGST @ 12% (highlighted)
        .Range("H26:J26").Merge
        .Range("H26").Value = "IGST :"
        .Range("H26").Font.Bold = True
        .Range("H26").Font.Size = 11 ' Increased font size
        .Range("H26").HorizontalAlignment = xlLeft
        .Range("H26").Interior.Color = RGB(255, 255, 200)
        .Range("H26").Font.Color = RGB(26, 26, 26)

        .Range("K26").Value = ""
        .Range("K26").Font.Bold = True
        .Range("K26").HorizontalAlignment = xlRight
        .Range("K26").Interior.Color = RGB(255, 255, 200)

        ' Row 27: CESS @ 0%
        .Range("H27:J27").Merge
        .Range("H27").Value = "CESS :"
        .Range("H27").Font.Bold = True
        .Range("H27").Font.Size = 11 ' Increased font size
        .Range("H27").HorizontalAlignment = xlLeft
        .Range("H27").Interior.Color = RGB(245, 245, 245)
        .Range("H27").Font.Color = RGB(26, 26, 26)

        .Range("K27").Value = ""
        .Range("K27").Font.Bold = True
        .Range("K27").HorizontalAlignment = xlRight
        .Range("K27").Interior.Color = RGB(216, 222, 233)

        ' Row 28: Total Tax (highlighted)
        .Range("H28:J28").Merge
        With .Range("H28")
            .Value = "Total Tax:"
            .Font.Bold = True
            .Font.Size = 11 ' Increased font size
            .Interior.Color = RGB(240, 240, 240)
            .Font.Color = RGB(26, 26, 26)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With

        .Range("K28").Value = ""
        .Range("K28").Font.Bold = True
        .Range("K28").HorizontalAlignment = xlRight
        .Range("K28").Interior.Color = RGB(240, 240, 240)

        ' Rows 29-30: Total Amount After Tax (merged across 2 rows for enhanced prominence)
        .Range("H29:J30").Merge
        .Range("H29").Value = "Total Amount After Tax:"
        .Range("H29").Font.Bold = True
        .Range("H29").Font.Size = 11 ' Increased font size
        .Range("H29").HorizontalAlignment = xlLeft
        .Range("H29").VerticalAlignment = xlCenter
        .Range("H29").Interior.Color = RGB(255, 255, 180)
        .Range("H29").Font.Color = RGB(26, 26, 26)

        .Range("K29:K30").Merge
        .Range("K29").Value = ""
        .Range("K29").Font.Bold = True
        .Range("K29").Font.Size = 11 ' Increased font size
        .Range("K29").HorizontalAlignment = xlRight
        .Range("K29").VerticalAlignment = xlCenter
        .Range("K29").Interior.Color = RGB(255, 255, 180)

        ' Set row heights for the merged final tax row
        .Rows(29).RowHeight = 18
        .Rows(30).RowHeight = 18

        ' Apply borders to entire tax summary section with professional color
        .Range("H26:K30").Borders.LineStyle = xlContinuous
        .Range("H26:K30").Borders.Color = RGB(204, 204, 204)

        ' Setup automatic tax calculation formulas for summary section
        Call UpdateMultiItemTaxCalculations(ws)

        ' Amount in words conversion integrated into signature section

        On Error GoTo 0

        ' --- Signature Section with Merged Cells ---
        On Error Resume Next

        ' Row 31: Signature headers with merged cells
        .Range("A31:C31").Merge
        .Range("A31").Value = "Transporter"
        .Range("A31").Font.Bold = True
        .Range("A31").HorizontalAlignment = xlCenter
        .Range("A31").Interior.Color = RGB(220, 220, 220)

        .Range("D31:G31").Merge
        .Range("D31").Value = "Receiver"
        .Range("D31").Font.Bold = True
        .Range("D31").HorizontalAlignment = xlCenter
        .Range("D31").Interior.Color = RGB(220, 220, 220)

        .Range("H31:K31").Merge
        .Range("H31").Value = "Certified that the particulars given above are true and correct"
        .Range("H31").Font.Bold = True
        .Range("H31").Font.Size = 10 ' Increased font size
        .Range("H31").HorizontalAlignment = xlCenter
        .Range("H31").VerticalAlignment = xlCenter
        .Range("H31").WrapText = True
        .Range("H31").Interior.Color = RGB(220, 220, 220)

        ' Rows 32-33: Mobile Number Section (merged across 2 rows)
        .Range("A32:C33").Merge
        .Range("A32").Value = "Mobile No: ___________________"
        .Range("A32").Font.Size = 10 ' Increased font size
        .Range("A32").HorizontalAlignment = xlCenter
        .Range("A32").VerticalAlignment = xlCenter
        .Range("A32").Interior.Color = RGB(250, 250, 250)

        .Range("D32:G33").Merge
        .Range("D32").Value = "Mobile No: ___________________"
        .Range("D32").Font.Size = 10 ' Increased font size
        .Range("D32").HorizontalAlignment = xlCenter
        .Range("D32").VerticalAlignment = xlCenter
        .Range("D32").Interior.Color = RGB(250, 250, 250)

        .Range("H32:K33").Merge
        .Range("H32").Value = "Mobile No: ___________________"
        .Range("H32").Font.Size = 10 ' Increased font size
        .Range("H32").HorizontalAlignment = xlCenter
        .Range("H32").VerticalAlignment = xlCenter
        .Range("H32").Interior.Color = RGB(250, 250, 250)

        ' Rows 34-36: Signature Space Section (merged across 3 rows)
        .Range("A34:C36").Merge
        .Range("A34").Value = ""
        .Range("A34").Interior.Color = RGB(250, 250, 250)

        .Range("D34:G36").Merge
        .Range("D34").Value = ""
        .Range("D34").Interior.Color = RGB(250, 250, 250)

        .Range("H34:K36").Merge
        .Range("H34").Value = ""
        .Range("H34").Interior.Color = RGB(250, 250, 250)

        ' Row 37: Signature Labels
        .Range("A37:C37").Merge
        .Range("A37").Value = "Transporter's Signature"
        .Range("A37").Font.Bold = True
        .Range("A37").Font.Size = 10 ' Increased font size
        .Range("A37").HorizontalAlignment = xlCenter
        .Range("A37").Interior.Color = RGB(211, 211, 211)

        .Range("D37:G37").Merge
        .Range("D37").Value = "Receiver's Signature"
        .Range("D37").Font.Bold = True
        .Range("D37").Font.Size = 10 ' Increased font size
        .Range("D37").HorizontalAlignment = xlCenter
        .Range("D37").Interior.Color = RGB(211, 211, 211)

        .Range("H37:K37").Merge
        .Range("H37").Value = "Authorized Signatory"
        .Range("H37").Font.Bold = True
        .Range("H37").Font.Size = 10 ' Increased font size
        .Range("H37").HorizontalAlignment = xlCenter
        .Range("H37").Interior.Color = RGB(211, 211, 211)

        ' Apply borders to entire signature section with professional color
        .Range("A31:K37").Borders.LineStyle = xlContinuous
        .Range("A31:K37").Borders.Color = RGB(204, 204, 204)

        ' Set specific row height for row 31 to accommodate wrapped text
        .Rows(31).RowHeight = 35

        ' Set standard row height for remaining signature rows
        For i = 32 To 37
            .Rows(i).RowHeight = 25 ' Increased height
        Next i
        .Rows(37).RowHeight = 31
        On Error GoTo 0
    End With

    ' --- Final Formatting ---
    With ws
        On Error Resume Next
        ' Font settings moved to beginning of code to avoid overriding header fonts
        On Error GoTo 0

        ' Apply professional page setup
        On Error Resume Next
        With .PageSetup
            .Orientation = xlPortrait
            .Zoom = False ' Let Excel handle scaling
            .FitToPagesWide = 1
            .FitToPagesTall = 1 ' Fit to one page vertically
            .LeftMargin = Application.InchesToPoints(0.3)
            .RightMargin = Application.InchesToPoints(0.3)
            .TopMargin = Application.InchesToPoints(0.3)
            .BottomMargin = Application.InchesToPoints(0.3)
            .HeaderMargin = Application.InchesToPoints(0.2)
            .FooterMargin = Application.InchesToPoints(0.2)
            .CenterHorizontally = True
            .CenterVertically = False
        End With
        On Error GoTo 0

        ' Add a subtle border around the entire invoice
        On Error Resume Next
        With .Range("A1:K37")
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThick
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThick
        End With
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

    ' Restore Excel alerts
    Application.DisplayAlerts = True

    MsgBox "GST TAX-INVOICE created successfully with professional buttons and auto-populated fields!", vbInformation
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
    ' Set up change monitoring for customer dropdown auto-population
    ' Since we're in a module, we'll use a different approach
    On Error Resume Next

    ' Note: Proper header "Details of Receiver (Billed to)" is created by CreateHeaderRow function
    ' No additional placeholder text needed - header is already properly formatted

    On Error GoTo 0
End Sub

Private Sub SetupStateCodeChangeEvents(ws As Worksheet)
    ' Simple state code setup - no automatic extraction needed
    ' State code dropdowns will show simple numeric codes only
    On Error GoTo 0
End Sub

Private Sub AutoFillConsigneeFromReceiver(ws As Worksheet)
    ' Automatically copy all receiver data to consignee fields
    On Error GoTo ErrorHandler

    With ws
        ' Copy Name from Receiver (C12:F12) to Consignee (I12:K12)
        .Range("I12").Value = .Range("C12").Value

        ' Copy Address from Receiver (C13:F13) to Consignee (I13:K13)
        .Range("I13").Value = .Range("C13").Value

        ' Copy GSTIN from Receiver (C14:F14) to Consignee (I14:K14)
        .Range("I14").Value = .Range("C14").Value

        ' Copy State from Receiver (C15:F15) to Consignee (I15:K15)
        .Range("I15").Value = .Range("C15").Value

        ' State code for consignee is now handled by a VLOOKUP formula in cell I16.
        ' This line is no longer needed as copying the state name to I15 will trigger the formula.

        ' Format consignee fields for manual editing (use default black font)
        .Range("I12:K12").Font.Color = RGB(26, 26, 26)  ' Standard black font
        .Range("I13:K13").Font.Color = RGB(26, 26, 26)  ' Standard black font
        .Range("I14:K14").Font.Color = RGB(26, 26, 26)  ' Standard black font
        .Range("I15:K15").Font.Color = RGB(26, 26, 26)  ' Standard black font
        .Range("I16:K16").Font.Color = RGB(26, 26, 26)  ' Standard black font
    End With

    Exit Sub

ErrorHandler:
    ' If auto-fill fails, continue silently
    On Error GoTo 0
End Sub

' 🧮 TAX CALCULATION FUNCTIONS

Public Sub SetupTaxCalculationFormulas(ws As Worksheet)
    ' Set up formulas for automatic tax calculations in the item table
    On Error Resume Next

    With ws
        ' For row 18 (first item row), set up formulas
        ' Column G (Amount) = Quantity * Rate
        .Range("G18").Formula = "=IF(AND(D18<>"""",F18<>""""),D18*F18,"""")"

        ' Column H (Taxable Value) = Amount (same as Amount for simplicity)
        .Range("H18").Formula = "=IF(G18<>"""",G18,"""")"

        ' Column I (IGST Rate) - VLOOKUP formula to get tax rate from HSN data
        .Range("I18").Formula = "=VLOOKUP(C18, warehouse!A:E, 5, FALSE)"

        ' Column J (IGST Amount) = Taxable Value * IGST Rate / 100
        .Range("J18").Formula = "=IF(AND(H18<>"""",I18<>""""),H18*I18/100,"""")"

        ' Column K (Total Amount) = Taxable Value + IGST Amount
        .Range("K18").Formula = "=IF(AND(H18<>"""",J18<>""""),H18+J18,"""")"

        ' Format the formula cells
        .Range("G18:K18").NumberFormat = "0.00"
        .Range("I18").NumberFormat = "0.00"
    End With

    On Error GoTo 0
End Sub

Public Sub UpdateMultiItemTaxCalculations(ws As Worksheet)
    ' Update tax calculations to sum all item rows
    On Error Resume Next

    With ws
        ' Row 22: Total Quantity
        .Range("D22").Formula = "=SUM(D18:D21)"
        .Range("D22").NumberFormat = "#,##0.00"

        ' Row 22: Sub Total calculations
        .Range("G22").Formula = "=SUM(G18:G21)"  ' Amount column
        .Range("H22").Formula = "=SUM(H18:H21)"  ' Taxable Value column
        .Range("G22:H22").NumberFormat = "#,##0.00"
        
        ' Row 22: IGST and Total Amount
        .Range("I22").Formula = "=SUM(J18:J21)" ' IGST Amount column
        .Range("K22").Formula = "=SUM(K18:K21)" ' Total Amount column
        .Range("I22:K22").NumberFormat = "#,##0.00"

        ' Row 23: Total Amount Before Tax
        .Range("K23").Formula = "=SUM(H18:H21)"

        ' Row 24: CGST (0 for interstate)
        .Range("K24").Value = 0

        ' Row 25: SGST (0 for interstate)
        .Range("K25").Value = 0

        ' Row 26: IGST
        .Range("K26").Formula = "=SUM(J18:J21)"

        ' Row 27: CESS (0 by default)
        .Range("K27").Value = 0

        ' Row 28: Total Tax
        .Range("K28").Formula = "=K24+K25+K26+K27"

        ' Row 29-30: Total Amount After Tax
        .Range("K29").Formula = "=K23+K28"
        .Range("K30").Formula = "=K29"

        ' Format all calculation cells
        .Range("K23:K30").NumberFormat = "#,##0.00"

        ' Update Amount in Words (A24:G25 merged cell)
        .Range("A24").Formula = "=NumberToWords(K29)"
    End With

    On Error GoTo 0
End Sub