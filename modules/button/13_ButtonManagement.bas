Option Explicit
' ===============================================================================
' MODULE: ButtonManagement
' DESCRIPTION: Functions for creating, managing, and removing buttons on the invoice worksheet
' ===============================================================================

Public Sub CreateInvoiceButtons(ws As Worksheet)
    ' Create professional buttons for invoice operations - Robust individual approach
    On Error GoTo ErrorHandler

    ' Remove any existing buttons first
    Call RemoveExistingButtons(ws)

    ' Add a small delay to ensure the worksheet is ready for button creation (macOS compatible)
    Dim startTime As Double
    startTime = Timer
    Do While Timer < startTime + 1  ' 1 second delay
        DoEvents  ' Allow system to process events
    Loop
    
    ' Create buttons with cell-based positioning for better visibility - MOVED TO COLUMNS R-U
    Call CreateButtonAtCell(ws, "R7", "Save Customer to Warehouse", "AddCustomerToWarehouseButton")
    Call CreateButtonAtCell(ws, "R9", "Save Invoice Record", "SaveInvoiceButton")
    Call CreateButtonAtCell(ws, "R11", "New Invoice", "NewInvoiceButton")
    Call CreateButtonAtCell(ws, "R13", "🔄 Refresh All", "RefreshButton")
    ' REMOVED: "Add New Item Row" button - functionality no longer needed
    Call CreateButtonAtCell(ws, "R19", "Export as PDF", "PrintAsPDFButton")
    Call CreateButtonAtCell(ws, "R21", "Print Invoice", "PrintButton")

    ' Add section headers
    Call CreateSectionHeaders(ws)

    ' Button creation completed successfully

    Exit Sub

ErrorHandler:
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

    ' Creating button at specified position

    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)

    If Err.Number = 0 And Not btn Is Nothing Then
        btn.Caption = caption
        btn.OnAction = onAction
        btn.Font.Name = "Segoe UI"
        btn.Font.Size = 9
        btn.Font.Bold = True

        ' Ensure button is on top
        btn.BringToFront

        ' Yield execution to allow Excel to process events
        DoEvents
    End If

    Err.Clear
    On Error GoTo 0
End Sub

Private Sub CreateSectionHeaders(ws As Worksheet)
    ' Create section headers AFTER the buttons for better organization
    On Error Resume Next

    ' INVOICE OPERATIONS header - MOVED TO COLUMN S (after buttons)
    ws.Range("S6").Value = "INVOICE OPERATIONS"
    ws.Range("S6").Font.Bold = True
    ws.Range("S6").Font.Size = 11
    ws.Range("S6").Font.Color = RGB(47, 80, 97)
    ws.Range("S6").HorizontalAlignment = xlLeft

    ' ITEM MANAGEMENT header - MOVED TO COLUMN S (after buttons)
    ws.Range("S14").Value = "ITEM MANAGEMENT"
    ws.Range("S14").Font.Bold = True
    ws.Range("S14").Font.Size = 11
    ws.Range("S14").Font.Color = RGB(47, 80, 97)
    ws.Range("S14").HorizontalAlignment = xlLeft

    ' PRINT & EXPORT header - MOVED TO COLUMN S (after buttons)
    ws.Range("S20").Value = "PRINT & EXPORT"
    ws.Range("S20").Font.Bold = True
    ws.Range("S20").Font.Size = 11
    ws.Range("S20").Font.Color = RGB(47, 80, 97)
    ws.Range("S20").HorizontalAlignment = xlLeft

    ' Footer note - MOVED TO COLUMN S (after buttons)
    ws.Range("S25").Value = "Click buttons for quick operations"
    ws.Range("S25").Font.Size = 8
    ws.Range("S25").Font.Italic = True
    ws.Range("S25").Font.Color = RGB(100, 100, 100)
    ws.Range("S25").HorizontalAlignment = xlLeft

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
