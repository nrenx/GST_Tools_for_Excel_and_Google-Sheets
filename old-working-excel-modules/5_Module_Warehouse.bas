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
            Array("4401", "Fuel wood, firewood, sawdust, wood waste and scrap", 2.5, 2.5, 5), _
            Array("4402", "Wood charcoal", 2.5, 2.5, 5), _
            Array("4403", "Wood in the rough (logs, unprocessed timber)", 9, 9, 18), _
            Array("4404", "Split poles, pickets, sticks, hoopwood, etc.", 6, 6, 12), _
            Array("4405", "Wood flour and wood wool", 6, 6, 12), _
            Array("4406", "Wooden railway or tramway sleepers", 6, 6, 12), _
            Array("4407", "Wood sawn or chipped", 9, 9, 18), _
            Array("4408", "Veneered wood and wood continuously shaped", 9, 9, 18), _
            Array("4409", "Moulded wood, flooring strips", 9, 9, 18), _
            Array("4410", "Particle board, oriented strand board (OSB), similar boards", 9, 9, 18), _
            Array("4412", "Plywood, veneered panels, laminated wood", 9, 9, 18), _
            Array("4413", "Densified wood", 9, 9, 18), _
            Array("4414", "Wooden frames for mirrors, photos, paintings", 9, 9, 18), _
            Array("4416", "Wooden barrels, casks, and other cooper’s products", 6, 6, 12), _
            Array("4417", "Wooden tools, tool handles, broom handles", 6, 6, 12), _
            Array("4418", "Builders’ joinery and carpentry of wood (doors, windows, etc.)", 9, 9, 18) _
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

        stateList = Array("Jammu and Kashmir", "Himachal Pradesh", "Punjab", "Chandigarh", "Uttarakhand", "Haryana", "Delhi", "Rajasthan", "Uttar Pradesh", "Bihar", "Sikkim", "Arunachal Pradesh", "Nagaland", "Manipur", "Mizoram", "Tripura", "Meghalaya", "Assam", "West Bengal", "Jharkhand", "Odisha", "Chhattisgarh", "Madhya Pradesh", "Gujarat", "Dadra and Nagar Haveli and Daman and Diu (merged)", "Maharashtra", "Karnataka", "Goa", "Lakshadweep", "Kerala", "Tamil Nadu", "Puducherry", "Andaman and Nicobar Islands", "Telangana", "Andhra Pradesh", "Ladakh")
        For i = 0 To UBound(stateList)
            .Cells(i + 2, 10).Value = stateList(i)
        Next i

        ' State Code List (Column K)
        .Range("K1").Value = "State_Code_List"
        .Range("K1").Font.Bold = True
        .Range("K1").Interior.Color = RGB(47, 80, 97)
        .Range("K1").Font.Color = RGB(255, 255, 255)

        stateCodeList = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "26", "27", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38")
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

        ' GST Type List (Column X)
        .Range("X1").Value = "GST_Type"
        .Range("X1").Font.Bold = True
        .Range("X1").Interior.Color = RGB(47, 80, 97)
        .Range("X1").Font.Color = RGB(255, 255, 255)
        .Range("X2").Value = "UNREGISTERED"
        
        ' Description List (Column Z)
        .Range("Z1").Value = "Description"
        .Range("Z1").Font.Bold = True
        .Range("Z1").Interior.Color = RGB(47, 80, 97)
        .Range("Z1").Font.Color = RGB(255, 255, 255)
        .Range("Z2").Value = "Casurina Wood"

        ' Increase column widths for customer data
        .Columns("M:T").ColumnWidth = 25

        ' Format customer headers
        .Range("M1:T1").Font.Bold = True
        .Range("M1:T1").Interior.Color = RGB(47, 80, 97)
        .Range("M1:T1").Font.Color = RGB(255, 255, 255)
        .Range("M1:T1").HorizontalAlignment = xlCenter

        ' Add sample customer data (simplified structure)
        ' Customer data is intentionally left blank for the user to populate.

        ' Auto-fit columns for other sections
        .Columns("A:L").AutoFit

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
        .Range("M1:T1").Borders.LineStyle = xlContinuous
        .Range("M1:T1").Borders.Color = RGB(204, 204, 204)
    End With
End Sub

' Note: The SetupDataValidation subroutine has been moved to Module_InvoiceEvents
' to better align with its role in managing UI interactions on the invoice sheet.

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
            Formula1:="=warehouse!$A$2:$A$17"  ' HSN codes from column A
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
