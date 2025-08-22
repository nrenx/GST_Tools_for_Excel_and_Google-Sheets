Option Explicit
' ===============================================================================
' MODULE: 18_DynamicTaxDisplay
' DESCRIPTION: Handles dynamic tax field display based on sale type, including
'              Interstate/Intrastate tax column switching and visibility.
' ===============================================================================

' ████████████████████████████████████████████████████████████████████████████████
' 🎯 DYNAMIC TAX DISPLAY MANAGEMENT
' ████████████████████████████████████████████████████████████████████████████████

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

Public Sub RefreshTaxDisplayForCurrentSaleType()
    ' Manual refresh function for tax display based on current sale type
    Dim ws As Worksheet
    Dim saleType As String
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("GST_Tax_Invoice_for_interstate")
    saleType = Trim(ws.Range("N7").Value)
    
    If saleType = "Interstate" Or saleType = "Intrastate" Then
        Call UpdateTaxFieldsDisplay(ws, saleType)
        ws.Calculate
        MsgBox "Tax display refreshed for " & saleType & " sale type!", vbInformation, "Tax Display Updated"
    Else
        MsgBox "Please select either 'Interstate' or 'Intrastate' in cell N7.", vbExclamation, "Invalid Sale Type"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error refreshing tax display: " & Err.Description, vbCritical, "Error"
End Sub
