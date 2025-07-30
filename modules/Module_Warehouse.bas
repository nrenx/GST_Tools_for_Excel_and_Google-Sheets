Option Explicit
' ===============================================================================
' MODULE: Module_Warehouse
' DESCRIPTION: Handles all data management related to the 'warehouse' sheet,
'              including customer data, HSN codes, and dropdown list setup.
' ===============================================================================

' ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
' 📋 WORKSHEET CREATION & DATA VALIDATION
' ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

Public Sub CreateWarehouseSheet()
    Dim ws As Worksheet
    Dim i As Integer, j As Integer
    Dim hsnData As Variant
    Dim uomList As Variant
    Dim transportList As Variant
    Dim stateList As Variant
    Dim stateCodeList As Variant
    Dim customerData As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("warehouse")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "warehouse"

    With ws
        ' ===== SECTION 1: HSN/SAC DATA (Columns A-E) =====
        ' HSN headers
        .Range("A1").Value = "HSN_Code"
        .Range("B1").Value = "Description"
        .Range("C1").Value = "CGST_Rate"
        .Range("D1").Value = "SGST_Rate"
        .Range("E1").Value = "IGST_Rate"

        ' Format HSN headers
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(47, 80, 97)
        .Range("A1:E1").Font.Color = RGB(255, 255, 255)
        .Range("A1:E1").HorizontalAlignment = xlCenter

        ' Add sample HSN data
        hsnData = Array( _
            Array("4403", "Casuarina Wood", 6, 6, 12), _
            Array("4407", "Sawn Wood", 6, 6, 12), _
            Array("4409", "Wood Flooring", 9, 9, 18), _
            Array("2501", "Salt", 2.5, 2.5, 5), _
            Array("1006", "Rice", 2.5, 2.5, 5), _
            Array("7208", "Steel Sheets", 9, 9, 18), _
            Array("8471", "Computer", 9, 9, 18), _
            Array("8517", "Mobile Phone", 9, 9, 18) _
        )

        For i = 0 To UBound(hsnData)
            For j = 0 To UBound(hsnData(i))
                .Cells(i + 2, j + 1).Value = hsnData(i)(j)  ' Starting at row 2, column A (1)
            Next j
        Next i

        ' ===== SECTION 2: VALIDATION LISTS =====
        ' UOM List (Column G)
        .Range("G1").Value = "UOM_List"
        .Range("G1").Font.Bold = True
        .Range("G1").Interior.Color = RGB(47, 80, 97)
        .Range("G1").Font.Color = RGB(255, 255, 255)

        uomList = Array("NOS", "KG", "MT", "CBM", "SQM", "LTR", "PCS", "BOX", "SET", "PAIR")
        For i = 0 To UBound(uomList)
            .Cells(i + 2, 7).Value = uomList(i)
        Next i

        ' Transport Mode List (Column H)
        .Range("H1").Value = "Transport_Mode_List"
        .Range("H1").Font.Bold = True
        .Range("H1").Interior.Color = RGB(47, 80, 97)
        .Range("H1").Font.Color = RGB(255, 255, 255)

        transportList = Array("By Lorry", "By Train", "By Air", "By Ship", "By Hand", "Courier", "Self Transport")
        For i = 0 To UBound(transportList)
            .Cells(i + 2, 8).Value = transportList(i)
        Next i

        ' State List (Column J)
        .Range("J1").Value = "State_List"
        .Range("J1").Font.Bold = True
        .Range("J1").Interior.Color = RGB(47, 80, 97)
        .Range("J1").Font.Color = RGB(255, 255, 255)

        stateList = Array("Andhra Pradesh", "Telangana", "Karnataka", "Tamil Nadu", "Kerala", "Maharashtra", "Gujarat", "Rajasthan", "Delhi", "Punjab")
        For i = 0 To UBound(stateList)
            .Cells(i + 2, 10).Value = stateList(i)
        Next i

        ' State Code List (Column K)
        .Range("K1").Value = "State_Code_List"
        .Range("K1").Font.Bold = True
        .Range("K1").Interior.Color = RGB(47, 80, 97)
        .Range("K1").Font.Color = RGB(255, 255, 255)

        stateCodeList = Array("37", "36", "29", "33", "32", "27", "24", "08", "07", "03")
        For i = 0 To UBound(stateCodeList)
            .Cells(i + 2, 11).Value = stateCodeList(i)
        Next i

        ' ===== SECTION 3: CUSTOMER MASTER DATA (Columns M-T) =====
        ' Customer headers (restored to original positions)
        .Range("M1").Value = "Customer_Name"
        .Range("N1").Value = "Address_Line1"
        .Range("O1").Value = "State"
        .Range("P1").Value = "State_Code"
        .Range("Q1").Value = "GSTIN"
        .Range("R1").Value = "Phone"
        .Range("S1").Value = "Email"
        .Range("T1").Value = "Contact_Person"

        ' Format customer headers
        .Range("M1:T1").Font.Bold = True
        .Range("M1:T1").Interior.Color = RGB(47, 80, 97)
        .Range("M1:T1").Font.Color = RGB(255, 255, 255)
        .Range("M1:T1").HorizontalAlignment = xlCenter

        ' Add sample customer data (simplified structure)
        customerData = Array( _
            Array("ABC Industries Ltd", "123 Industrial Area, Sector 15, Tirupati", "Andhra Pradesh", "37", "37ABCDE1234F1Z5", "9876543210", "abc@industries.com", "Mr. Sharma"), _
            Array("XYZ Trading Co", "456 Market Street, Near Bus Stand, Vijayawada", "Andhra Pradesh", "37", "37XYZAB5678G2H6", "9876543211", "xyz@trading.com", "Ms. Patel"), _
            Array("PQR Enterprises", "789 Commercial Complex, Phase 2, Visakhapatnam", "Andhra Pradesh", "37", "37PQRST9012I3J7", "9876543212", "pqr@enterprises.com", "Mr. Kumar") _
        )

        For i = 0 To UBound(customerData)
            For j = 0 To UBound(customerData(i))
                .Cells(i + 2, j + 13).Value = customerData(i)(j)  ' Starting at column M (13)
            Next j
        Next i

        ' Auto-fit columns
        .Columns.AutoFit

        ' Add borders to all sections
        ' HSN data borders
        .Range("A1:E" & UBound(hsnData) + 2).Borders.LineStyle = xlContinuous
        .Range("A1:E" & UBound(hsnData) + 2).Borders.Color = RGB(204, 204, 204)

        ' Validation lists borders
        .Range("G1:G" & UBound(uomList) + 2).Borders.LineStyle = xlContinuous
        .Range("H1:H" & UBound(transportList) + 2).Borders.LineStyle = xlContinuous
        .Range("J1:J" & UBound(stateList) + 2).Borders.LineStyle = xlContinuous
        .Range("K1:K" & UBound(stateCodeList) + 2).Borders.LineStyle = xlContinuous

        ' Customer data borders
        .Range("M1:T" & UBound(customerData) + 2).Borders.LineStyle = xlContinuous
        .Range("M1:T" & UBound(customerData) + 2).Borders.Color = RGB(204, 204, 204)
    End With
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
        ' Allow both dropdown selection AND manual text entry
        .Range("E18:E21").Validation.Delete
        .Range("E18:E21").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$G$2:$G$11"  ' UOM list from column G
        .Range("E18:E21").Validation.IgnoreBlank = True
        .Range("E18:E21").Validation.InCellDropdown = True
        .Range("E18:E21").Validation.ShowError = False  ' Allow manual text entry
        .Range("E18:E21").Font.Color = RGB(26, 26, 26)  ' Standard black font

        ' Transport Mode dropdown with manual text entry capability (F7)
        ' Allow both dropdown selection AND manual text entry
        .Range("F7").Validation.Delete
        .Range("F7").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$H$2:$H$8"  ' Transport modes from column H
        .Range("F7").Validation.IgnoreBlank = True
        .Range("F7").Validation.InCellDropdown = True
        .Range("F7").Validation.ShowError = False  ' Allow manual text entry

        ' Set default transport mode
        If .Range("F7").Value = "" Then
            .Range("F7").Value = "By Lorry"
        End If

        ' State dropdown for Receiver (Row 15, Column C15:F15)
        .Range("C15").Validation.Delete
        .Range("C15").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$J$2:$J$11"  ' State list from column J
        .Range("C15").Validation.IgnoreBlank = True
        .Range("C15").Validation.InCellDropdown = True
        .Range("C15").Validation.ShowError = False  ' Allow manual text entry
        .Range("C15").Font.Color = RGB(26, 26, 26)  ' Standard black font

        ' State dropdown for Consignee (Row 15, Column I15:K15)
        .Range("I15").Validation.Delete
        .Range("I15").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$J$2:$J$11"  ' State list from column J
        .Range("I15").Validation.IgnoreBlank = True
        .Range("I15").Validation.InCellDropdown = True
        .Range("I15").Validation.ShowError = False  ' Allow manual text entry
        .Range("I15").Font.Color = RGB(26, 26, 26)  ' Standard black font

        ' State Code dropdown for Receiver (Row 16, Column C16) - shows simple numeric codes
        .Range("C16").Validation.Delete
        .Range("C16").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$K$2:$K$11"  ' Simple state code list from column K
        .Range("C16").Validation.IgnoreBlank = True
        .Range("C16").Validation.InCellDropdown = True
        .Range("C16").Validation.ShowError = False  ' Allow manual text entry
        .Range("C16").Font.Color = RGB(26, 26, 26)  ' Standard black font

        ' State Code dropdown for Consignee (Row 16, Column I16) - shows simple numeric codes
        .Range("I16").Validation.Delete
        .Range("I16").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$K$2:$K$11"  ' Simple state code list from column K
        .Range("I16").Validation.IgnoreBlank = True
        .Range("I16").Validation.InCellDropdown = True
        .Range("I16").Validation.ShowError = False  ' Allow manual text entry
        .Range("I16").Font.Color = RGB(26, 26, 26)  ' Standard black font

    End With

    On Error GoTo 0
End Sub

' ===== CUSTOMER DATABASE INTEGRATION =====

Public Sub SetupCustomerDropdown(ws As Worksheet)
    ' Setup customer dropdown and auto-population
    Dim dropdownWs As Worksheet
    On Error Resume Next

    ' Ensure supporting worksheets exist
    Call EnsureAllSupportingWorksheetsExist

    Set dropdownWs = GetOrCreateWorksheet("warehouse")

    If dropdownWs Is Nothing Then
        Exit Sub
    End If

    With ws
        ' Customer dropdown with manual text entry capability for Receiver (row 12, column C)
        ' Allow both dropdown selection AND manual text entry
        .Range("C12").Validation.Delete
        .Range("C12").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$M$2:$M$10"  ' Customer names (restored to column M)
        .Range("C12").Validation.IgnoreBlank = True
        .Range("C12").Validation.InCellDropdown = True
        .Range("C12").Validation.ShowError = False  ' Allow manual text entry
        .Range("C12").Font.Bold = True
        .Range("C12").Interior.Color = RGB(255, 255, 255)  ' White background
        .Range("C12").Font.Color = RGB(26, 26, 26)  ' Standard black font

        ' Customer dropdown with manual text entry capability for Consignee (row 12, column I)
        ' Allow both dropdown selection AND manual text entry
        .Range("I12").Validation.Delete
        .Range("I12").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$M$2:$M$10"  ' Customer names (restored to column M)
        .Range("I12").Validation.IgnoreBlank = True
        .Range("I12").Validation.InCellDropdown = True
        .Range("I12").Validation.ShowError = False  ' Allow manual text entry
        .Range("I12").Font.Bold = True
        .Range("I12").Interior.Color = RGB(255, 255, 255)  ' White background
        .Range("I12").Font.Color = RGB(26, 26, 26)  ' Standard black font

        ' Set fixed state code for Andhra Pradesh (no dropdown needed)
        .Range("C10").Validation.Delete  ' Remove any existing validation
        .Range("C10").Value = "37"  ' Fixed value for Andhra Pradesh
        .Range("C10").Font.Bold = True
        .Range("C10").Interior.Color = RGB(245, 245, 245)  ' Light grey background
        .Range("C10").Font.Color = RGB(26, 26, 26)  ' Dark text
        .Range("C10").HorizontalAlignment = xlLeft
    End With

    On Error GoTo 0
End Sub

Public Sub PopulateCustomerDetails(ws As Worksheet, customerName As String)
    ' Automatically populate customer details when customer is selected
    Dim customerDetails As Variant

    If customerName = "" Then Exit Sub

    customerDetails = GetCustomerDetails(customerName)

    With ws
        ' Populate customer details in the party details section
        .Range("C13").Value = customerDetails(2)  ' Address
        .Range("C14").Value = customerDetails(8)  ' GSTIN
        .Range("C15").Value = customerDetails(5)  ' State
        .Range("C16").Value = customerDetails(6)  ' State Code

        ' Keep state code fixed as "37" for Andhra Pradesh (don't override)
        .Range("C10").Value = "37"  ' Always keep as 37 for Andhra Pradesh
    End With
End Sub

Public Function GetCustomerDetails(customerName As String) As Variant
    ' Get customer details from warehouse sheet (Customer section - columns M-T)
    Dim dropdownWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim customerDetails(11) As String  ' Array to hold customer details

    On Error GoTo ErrorHandler

    ' Ensure supporting worksheets exist
    Call EnsureAllSupportingWorksheetsExist

    Set dropdownWs = GetOrCreateWorksheet("warehouse")

    If dropdownWs Is Nothing Then
        GetCustomerDetails = customerDetails
        Exit Function
    End If

    ' Customer data starts at row 2, column M (Customer_Name is in column M)
    lastRow = dropdownWs.Cells(dropdownWs.Rows.Count, 13).End(xlUp).Row

    For i = 2 To lastRow
        If UCase(dropdownWs.Cells(i, 13).Value) = UCase(customerName) Then
            ' Found the customer, populate details array (restored to original structure)
            customerDetails(0) = ""                             ' Customer_ID (not in structure)
            customerDetails(1) = dropdownWs.Cells(i, 13).Value  ' Customer_Name (Column M)
            customerDetails(2) = dropdownWs.Cells(i, 14).Value  ' Address_Line1 (Column N)
            customerDetails(3) = ""                             ' Address_Line2 (not in structure)
            customerDetails(4) = ""                             ' City (not in structure)
            customerDetails(5) = dropdownWs.Cells(i, 15).Value  ' State (Column O)
            customerDetails(6) = dropdownWs.Cells(i, 16).Value  ' State_Code (Column P)
            customerDetails(7) = ""                             ' PIN_Code (not in structure)
            customerDetails(8) = dropdownWs.Cells(i, 17).Value  ' GSTIN (Column Q)
            customerDetails(9) = dropdownWs.Cells(i, 18).Value  ' Phone (Column R)
            customerDetails(10) = dropdownWs.Cells(i, 19).Value ' Email (Column S)
            customerDetails(11) = dropdownWs.Cells(i, 20).Value ' Contact_Person (Column T)
            Exit For
        End If
    Next i

    GetCustomerDetails = customerDetails
    Exit Function

ErrorHandler:
    GetCustomerDetails = customerDetails
End Function

' ===== HSN/SAC CODE LOOKUP SYSTEM =====

Public Sub SetupHSNDropdown(ws As Worksheet)
    ' Setup HSN code dropdown for item rows
    Dim dropdownWs As Worksheet
    On Error Resume Next

    ' Ensure supporting worksheets exist
    Call EnsureAllSupportingWorksheetsExist

    Set dropdownWs = GetOrCreateWorksheet("warehouse")

    If dropdownWs Is Nothing Then
        Exit Sub
    End If

    With ws
        ' HSN Code dropdown with manual text entry capability (Column C: 18-21)
        ' Allow both dropdown selection AND manual text entry
        .Range("C18:C21").Validation.Delete
        .Range("C18:C21").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertInformation, _
            Formula1:="=warehouse!$A$2:$A$20"  ' HSN codes from column A
        .Range("C18:C21").Validation.IgnoreBlank = True
        .Range("C18:C21").Validation.InCellDropdown = True
        .Range("C18:C21").Validation.ShowError = False  ' Allow manual text entry
        .Range("C18:C21").Font.Color = RGB(26, 26, 26)  ' Standard black font
    End With

    On Error GoTo 0
End Sub

Public Function GetHSNDetails(hsnCode As String) As Variant
    ' Get HSN details from warehouse sheet (HSN section - columns A-E)
    Dim dropdownWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim hsnDetails(7) As String  ' Array to hold HSN details

    On Error GoTo ErrorHandler

    ' Ensure supporting worksheets exist
    Call EnsureAllSupportingWorksheetsExist

    Set dropdownWs = GetOrCreateWorksheet("warehouse")

    If dropdownWs Is Nothing Then
        GetHSNDetails = hsnDetails
        Exit Function
    End If

    ' HSN data starts at row 2, column A (HSN_Code is in column A)
    lastRow = dropdownWs.Cells(dropdownWs.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If UCase(dropdownWs.Cells(i, 1).Value) = UCase(hsnCode) Then
            ' Found the HSN code, get all details
            hsnDetails(0) = dropdownWs.Cells(i, 1).Value  ' HSN_Code (Column A)
            hsnDetails(1) = dropdownWs.Cells(i, 2).Value  ' Description (Column B)
            hsnDetails(2) = ""                            ' UOM (not in new structure)
            hsnDetails(3) = dropdownWs.Cells(i, 3).Value  ' CGST_Rate (Column C)
            hsnDetails(4) = dropdownWs.Cells(i, 4).Value  ' SGST_Rate (Column D)
            hsnDetails(5) = dropdownWs.Cells(i, 5).Value  ' IGST_Rate (Column E)
            hsnDetails(6) = ""                            ' CESS_Rate (not in new structure)
            hsnDetails(7) = ""                            ' Category (not in new structure)
            Exit For
        End If
    Next i

    GetHSNDetails = hsnDetails
    Exit Function

ErrorHandler:
    GetHSNDetails = hsnDetails
End Function

Public Sub PopulateHSNDetails(targetCell As Range)
    ' Auto-populate HSN details when HSN code is selected
    Dim ws As Worksheet
    Dim hsnCode As String
    Dim hsnDetails As Variant
    Dim rowNum As Long

    On Error GoTo ErrorHandler

    Set ws = targetCell.Worksheet
    If ws.Name <> "GST_Tax_Invoice_for_interstate" Then Exit Sub

    hsnCode = targetCell.Value
    If hsnCode = "" Then Exit Sub

    rowNum = targetCell.Row
    hsnDetails = GetHSNDetails(hsnCode)

    ' Populate description (Column B)
    If hsnDetails(1) <> "" Then
        ws.Cells(rowNum, 2).Value = hsnDetails(1)
    End If

    ' Populate UOM (Column E)
    If hsnDetails(2) <> "" Then
        ws.Cells(rowNum, 5).Value = hsnDetails(2)
    End If

    ' Populate IGST Rate (Column I) - for interstate sales
    If hsnDetails(5) <> "" Then
        ws.Cells(rowNum, 9).Value = hsnDetails(5)
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error populating HSN details: " & Err.Description, vbCritical
End Sub

Public Sub AddHSNToMaster()
    ' Add new HSN code to warehouse sheet (HSN section - columns A-E)
    Dim dropdownWs As Worksheet
    Dim lastRow As Long
    Dim hsnCode As String, description As String
    Dim cgstRate As String, sgstRate As String, igstRate As String

    On Error GoTo ErrorHandler

    ' Ensure supporting worksheets exist
    Call EnsureAllSupportingWorksheetsExist

    Set dropdownWs = GetOrCreateWorksheet("warehouse")

    If dropdownWs Is Nothing Then
        MsgBox "warehouse sheet not found!", vbExclamation
        Exit Sub
    End If

    ' HSN data starts at row 2, find last row in HSN section (column A)
    lastRow = dropdownWs.Cells(dropdownWs.Rows.Count, 1).End(xlUp).Row

    ' Simple input form (you can enhance this with a UserForm)
    hsnCode = InputBox("Enter HSN/SAC Code:", "Add New HSN Code")
    If hsnCode = "" Then Exit Sub

    description = InputBox("Enter Description:", "Add New HSN Code")
    cgstRate = InputBox("Enter CGST Rate (%):", "Add New HSN Code")
    sgstRate = InputBox("Enter SGST Rate (%):", "Add New HSN Code")
    igstRate = InputBox("Enter IGST Rate (%):", "Add New HSN Code")

    With dropdownWs
        ' Add HSN data to columns A-E (new structure)
        .Cells(lastRow + 1, 1).Value = hsnCode           ' Column A - HSN_Code
        .Cells(lastRow + 1, 2).Value = description       ' Column B - Description
        .Cells(lastRow + 1, 3).Value = Val(cgstRate)     ' Column C - CGST_Rate
        .Cells(lastRow + 1, 4).Value = Val(sgstRate)     ' Column D - SGST_Rate
        .Cells(lastRow + 1, 5).Value = Val(igstRate)     ' Column E - IGST_Rate

        ' Add borders
        .Range("A" & lastRow + 1 & ":E" & lastRow + 1).Borders.LineStyle = xlContinuous
        .Range("A" & lastRow + 1 & ":E" & lastRow + 1).Borders.Color = RGB(204, 204, 204)
    End With

    MsgBox "HSN code added successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error adding HSN code: " & Err.Description, vbCritical
End Sub