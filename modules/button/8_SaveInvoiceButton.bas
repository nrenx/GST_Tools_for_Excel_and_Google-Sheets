Option Explicit
' ===============================================================================
' MODULE: SaveInvoiceButton
' DESCRIPTION: Button function to save complete invoice record to Master sheet 
'              for future reference with GST compliance
' ===============================================================================

Public Sub SaveInvoiceButton()
    ' Button function: Save complete invoice record to Master sheet for future reference - UPDATED FOR NEW TAX FIELDS
    Dim invoiceWs As Worksheet
    Dim masterWs As Worksheet
    Dim invoiceNumber As String, invoiceDate As String, customerName As String
    Dim customerGSTIN As String, customerState As String, customerStateCode As String
    Dim hsnCodes As String, itemDescriptions As String, totalQuantity As String, uomList As String
    Dim taxableTotal As Double, igstTotal As Double, cgstTotal As Double, sgstTotal As Double, grandTotal As Double, totalQty As Double
    Dim i As Long
    Dim saleType As String, igstRate As String, cgstRate As String, sgstRate As String
    Dim lastRow As Long
    Dim response As VbMsgBoxResult
    On Error GoTo ErrorHandler

    Set invoiceWs = ThisWorkbook.Worksheets("GST_Tax_Invoice_for_interstate")
    Set masterWs = ThisWorkbook.Worksheets("Master")

    ' Get invoice details for GST compliance - UPDATED FOR NEW LAYOUT
    invoiceNumber = Trim(invoiceWs.Range("C7").Value)
    invoiceDate = Trim(invoiceWs.Range("C8").Value)
    customerName = Trim(invoiceWs.Range("C12").Value)
    customerGSTIN = Trim(invoiceWs.Range("C14").Value)
    customerState = Trim(invoiceWs.Range("C15").Value)
    customerStateCode = Trim(invoiceWs.Range("C16").Value)
    saleType = Trim(invoiceWs.Range("N7").Value)  ' NEW: Sale Type

    ' Calculate totals and collect item details from item table
    For i = 18 To 21 ' Check all possible item rows
        If invoiceWs.Cells(i, "H").Value <> "" And IsNumeric(invoiceWs.Cells(i, "H").Value) Then
            taxableTotal = taxableTotal + invoiceWs.Cells(i, "H").Value
        End If
        If invoiceWs.Cells(i, "J").Value <> "" And IsNumeric(invoiceWs.Cells(i, "J").Value) Then
            igstTotal = igstTotal + invoiceWs.Cells(i, "J").Value
        End If
        If invoiceWs.Cells(i, "L").Value <> "" And IsNumeric(invoiceWs.Cells(i, "L").Value) Then
            cgstTotal = cgstTotal + invoiceWs.Cells(i, "L").Value  ' NEW: CGST Amount
        End If
        If invoiceWs.Cells(i, "N").Value <> "" And IsNumeric(invoiceWs.Cells(i, "N").Value) Then
            sgstTotal = sgstTotal + invoiceWs.Cells(i, "N").Value  ' NEW: SGST Amount
        End If
        If invoiceWs.Cells(i, "O").Value <> "" And IsNumeric(invoiceWs.Cells(i, "O").Value) Then
            grandTotal = grandTotal + invoiceWs.Cells(i, "O").Value  ' NEW: Total Amount column O
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

    ' Calculate tax rates based on sale type
    If taxableTotal > 0 Then
        If saleType = "Interstate" Then
            igstRate = Format((igstTotal / taxableTotal) * 100, "0.00") & "%"
            cgstRate = "0.00%"
            sgstRate = "0.00%"
        ElseIf saleType = "Intrastate" Then
            igstRate = "0.00%"
            cgstRate = Format((cgstTotal / taxableTotal) * 100, "0.00") & "%"
            sgstRate = Format((sgstTotal / taxableTotal) * 100, "0.00") & "%"
        Else
            igstRate = "18.00%"
            cgstRate = "0.00%"
            sgstRate = "0.00%"
        End If
    Else
        igstRate = "18.00%"
        cgstRate = "0.00%"
        sgstRate = "0.00%"
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
    ' Save complete GST-compliant invoice data to Master sheet - UPDATED FOR NEW TAX FIELDS
    With masterWs
        .Cells(lastRow, "A").Value = invoiceNumber          ' Column A: Invoice_Number
        .Cells(lastRow, "B").Value = invoiceDate            ' Column B: Invoice_Date
        .Cells(lastRow, "C").Value = customerName           ' Column C: Customer_Name
        .Cells(lastRow, "D").Value = customerGSTIN          ' Column D: Customer_GSTIN
        .Cells(lastRow, "E").Value = customerState          ' Column E: Customer_State
        .Cells(lastRow, "F").Value = customerStateCode      ' Column F: Customer_State_Code
        .Cells(lastRow, "G").Value = taxableTotal           ' Column G: Total_Taxable_Value
        .Cells(lastRow, "H").Value = saleType               ' Column H: Sale_Type
        .Cells(lastRow, "I").Value = igstRate               ' Column I: IGST_Rate
        .Cells(lastRow, "J").Value = igstTotal              ' Column J: IGST_Amount
        .Cells(lastRow, "K").Value = cgstRate               ' Column K: CGST_Rate
        .Cells(lastRow, "L").Value = cgstTotal              ' Column L: CGST_Amount
        .Cells(lastRow, "M").Value = sgstRate               ' Column M: SGST_Rate
        .Cells(lastRow, "N").Value = sgstTotal              ' Column N: SGST_Amount
        .Cells(lastRow, "O").Value = igstTotal + cgstTotal + sgstTotal  ' Column O: Total_Tax_Amount
        .Cells(lastRow, "P").Value = grandTotal             ' Column P: Total_Invoice_Value
        .Cells(lastRow, "Q").Value = hsnCodes               ' Column Q: HSN_Codes
        .Cells(lastRow, "R").Value = itemDescriptions       ' Column R: Item_Description
        .Cells(lastRow, "S").Value = totalQty               ' Column S: Quantity
        .Cells(lastRow, "T").Value = uomList                ' Column T: UOM
        .Cells(lastRow, "U").Value = Now                    ' Column U: Date_Created

        ' Add borders for the new record - EXPANDED TO COLUMN U
        .Range("A" & lastRow & ":U" & lastRow).Borders.LineStyle = xlContinuous
        .Range("A" & lastRow & ":U" & lastRow).Borders.Color = RGB(204, 204, 204)
    End With

    MsgBox "Invoice " & invoiceNumber & " saved successfully to Master sheet!" & vbCrLf & _
           "Customer: " & customerName & vbCrLf & _
           "Total Taxable Value: " & ChrW(8377) & Format(taxableTotal, "#,##0.00") & vbCrLf & _
           "IGST Amount: " & ChrW(8377) & Format(igstTotal, "#,##0.00") & vbCrLf & _
           "Total Invoice Value: " & ChrW(8377) & Format(grandTotal, "#,##0.00") & vbCrLf & vbCrLf & _
           "Record saved for GST audit and return filing purposes.", vbInformation, "GST Invoice Record Saved"
    Exit Sub

ErrorHandler:
    MsgBox "Error saving invoice: " & Err.Description, vbCritical, "Error"
End Sub
