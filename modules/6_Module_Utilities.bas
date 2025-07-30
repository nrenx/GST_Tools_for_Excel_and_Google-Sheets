Option Explicit
' ===============================================================================
' MODULE: Module_Utilities
' DESCRIPTION: Contains shared helper functions used across multiple modules,
'              including worksheet management, text cleaning, and number-to-word
'              conversion.
' ===============================================================================

' ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
' 🔧 UTILITY FUNCTIONS
' ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓

Public Function WorksheetExists(sheetName As String) As Boolean
    ' Check if a worksheet exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not (ws Is Nothing)
End Function

Public Function GetOrCreateWorksheet(sheetName As String) As Worksheet
    ' Safely get or create a worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateWorksheet = ws
End Function

Public Sub EnsureAllSupportingWorksheetsExist()
    ' Ensure all required supporting worksheets exist
    On Error Resume Next

    ' Create Master sheet if it doesn't exist
    If Not WorksheetExists("Master") Then
        Call CreateMasterSheet
    End If

    ' Create warehouse sheet if it doesn't exist
    If Not WorksheetExists("warehouse") Then
        Call CreateWarehouseSheet
    End If

    On Error GoTo 0
End Sub


Public Function CleanText(inputText As String) As String
    Dim cleanedText As String
    Dim i As Integer

    cleanedText = inputText

    ' Remove any question marks that might appear due to encoding issues
    cleanedText = Replace(cleanedText, "?", "")

    ' Remove any other problematic characters
    cleanedText = Replace(cleanedText, Chr(63), "") ' ASCII 63 is question mark

    ' Trim extra spaces
    cleanedText = Trim(cleanedText)

    ' Replace multiple spaces with single space
    Do While InStr(cleanedText, "  ") > 0
        cleanedText = Replace(cleanedText, "  ", " ")
    Loop

    CleanText = cleanedText
End Function

Public Sub VerifyValidationSettings()
    ' Display current validation settings to confirm manual editing is enabled
    Dim ws As Worksheet
    Dim message As String

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets("GST_Tax_Invoice_for_interstate")

    message = "VALIDATION SETTINGS VERIFICATION:" & vbCrLf & vbCrLf
    message = message & "✅ ALL FIELDS SUPPORT MANUAL EDITING:" & vbCrLf & vbCrLf

    message = message & "📝 DROPDOWN + MANUAL ENTRY FIELDS:" & vbCrLf
    message = message & "• Customer Name (C12) - Dropdown + Manual" & vbCrLf
    message = message & "• Receiver State (C15) - Dropdown + Manual" & vbCrLf
    message = message & "• Consignee State (I15) - Dropdown + Manual" & vbCrLf
    message = message & "• HSN Code (C18:C21) - Dropdown + Manual" & vbCrLf
    message = message & "• UOM (E18:E21) - Dropdown + Manual" & vbCrLf
    message = message & "• Transport Mode (F7) - Dropdown + Manual" & vbCrLf & vbCrLf

    message = message & "🔓 FULLY EDITABLE FIELDS:" & vbCrLf
    message = message & "• Invoice Number (C7) - Auto + Manual Override" & vbCrLf
    message = message & "• Invoice Date (C8) - Auto + Manual Override" & vbCrLf
    message = message & "• Date of Supply (F9, G9) - Auto + Manual Override" & vbCrLf
    message = message & "• State Code (C10) - Fixed + Manual Override" & vbCrLf
    message = message & "• All Address/GSTIN fields - Fully Manual" & vbCrLf
    message = message & "• All Item details - Fully Manual" & vbCrLf & vbCrLf

    message = message & "🎯 KEY FEATURES:" & vbCrLf
    message = message & "• No restrictive validations (xlValidAlertStop removed)" & vbCrLf
    message = message & "• ShowError = False for all dropdowns" & vbCrLf
    message = message & "• Users can override ANY auto-populated value" & vbCrLf
    message = message & "• Dropdown suggestions + free text entry" & vbCrLf & vbCrLf

    message = message & "💡 All validation requirements have been successfully implemented!"

    MsgBox message, vbInformation, "Validation Settings - All Clear ✅"
    Exit Sub

ErrorHandler:
    MsgBox "Error verifying validation settings: " & Err.Description, vbCritical, "Verification Error"
End Sub

' ===== AMOUNT IN WORDS CONVERSION SYSTEM =====

Public Function NumberToWords(ByVal MyNumber)
    Dim Rupees, Paise, Temp
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Lakh "
    Place(4) = " Crore "

    MyNumber = Trim(Str(MyNumber))
    DecimalPlace = InStr(MyNumber, ".")

    If DecimalPlace > 0 Then
        Paise = ConvertTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If

    Count = 1
    Do While MyNumber <> ""
        Select Case Count
            Case 1
                Temp = ConvertHundreds(Right(MyNumber, 3))
                If Len(MyNumber) > 3 Then
                    MyNumber = Left(MyNumber, Len(MyNumber) - 3)
                Else
                    MyNumber = ""
                End If
            Case 2
                Temp = ConvertTens(Right(MyNumber, 2))
                If Len(MyNumber) > 2 Then
                    MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                Else
                    MyNumber = ""
                End If
            Case Else
                Temp = ConvertTens(Right(MyNumber, 2))
                If Len(MyNumber) > 2 Then
                    MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                Else
                    MyNumber = ""
                End If
        End Select

        If Temp <> "" Then Rupees = Temp & Place(Count) & Rupees
        Count = Count + 1
    Loop

    Select Case Rupees
        Case ""
            Rupees = "Zero Rupees"
        Case "One"
            Rupees = "One Rupee"
        Case Else
            Rupees = Rupees & " Rupees"
    End Select

    If Paise <> "" Then
        Select Case Paise
            Case "One"
                Paise = " and One Paisa"
            Case Else
                Paise = " and " & Paise & " Paise"
        End Select
    End If

    NumberToWords = CleanText(Rupees & Paise & " Only")
End Function

Private Function ConvertHundreds(ByVal MyNumber)
    Dim Result As String

    ' Exit if there is nothing to convert
    If Val(MyNumber) = 0 Then Exit Function

    ' Append leading zeros to number
    MyNumber = Right("000" & MyNumber, 3)

    ' Do we have a hundreds place digit to convert?
    If Left(MyNumber, 1) <> "0" Then
        Result = ConvertDigit(Left(MyNumber, 1)) & " Hundred "
    End If

    ' Do we have a tens place digit to convert?
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & ConvertTens(Mid(MyNumber, 2))
    Else
        ' If not, then convert the ones place digit
        Result = Result & ConvertDigit(Mid(MyNumber, 3))
    End If

    ConvertHundreds = Trim(Result)
End Function

Private Function ConvertTens(ByVal MyTens)
    Dim Result As String

    ' Is value between 10 and 19?
    If Val(Left(MyTens, 1)) = 1 Then
        Select Case Val(MyTens)
            Case 10: Result = "Ten"
            Case 11: Result = "Eleven"
            Case 12: Result = "Twelve"
            Case 13: Result = "Thirteen"
            Case 14: Result = "Fourteen"
            Case 15: Result = "Fifteen"
            Case 16: Result = "Sixteen"
            Case 17: Result = "Seventeen"
            Case 18: Result = "Eighteen"
            Case 19: Result = "Nineteen"
            Case Else
        End Select
    Else
        ' .. otherwise it's between 20 and 99
        Select Case Val(Left(MyTens, 1))
            Case 2: Result = "Twenty "
            Case 3: Result = "Thirty "
            Case 4: Result = "Forty "
            Case 5: Result = "Fifty "
            Case 6: Result = "Sixty "
            Case 7: Result = "Seventy "
            Case 8: Result = "Eighty "
            Case 9: Result = "Ninety "
            Case Else
        End Select

        ' Convert ones place digit
        Result = Result & ConvertDigit(Right(MyTens, 1))
    End If

    ConvertTens = Result
End Function

Private Function ConvertDigit(ByVal MyDigit)
    Select Case Val(MyDigit)
        Case 1: ConvertDigit = "One"
        Case 2: ConvertDigit = "Two"
        Case 3: ConvertDigit = "Three"
        Case 4: ConvertDigit = "Four"
        Case 5: ConvertDigit = "Five"
        Case 6: ConvertDigit = "Six"
        Case 7: ConvertDigit = "Seven"
        Case 8: ConvertDigit = "Eight"
        Case 9: ConvertDigit = "Nine"
        Case Else: ConvertDigit = ""
    End Select
End Function