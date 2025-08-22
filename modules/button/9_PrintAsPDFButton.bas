Option Explicit
' ===============================================================================
' MODULE: PrintAsPDFButton
' DESCRIPTION: Button function to export invoice as a two-page PDF (Original and Duplicate)
'              with enhanced macOS compatibility
' ===============================================================================

Public Sub PrintAsPDFButton()
    ' Button function: Export invoice as a two-page PDF (Original and Duplicate)
    Dim originalWs As Worksheet
    Dim duplicateWs As Worksheet
    Dim invoiceNumber As String
    Dim cleanInvoiceNumber As String
    Dim pdfPath As String
    Dim fullPath As String
    Dim fso As Object
    Dim cell As Range
    On Error GoTo ErrorHandler

    Set originalWs = ThisWorkbook.Worksheets("GST_Tax_Invoice_for_interstate")

    ' Ensure warehouse worksheet exists to prevent file dialog errors
    Call EnsureAllSupportingWorksheetsExist

    ' Get invoice number for filename
    invoiceNumber = Trim(originalWs.Range("C7").Value)

    If invoiceNumber = "" Then
        MsgBox "Please ensure invoice number is filled before exporting to PDF.", vbExclamation, "Missing Invoice Number"
        Exit Sub
    End If

    ' Clean invoice number for filename
    cleanInvoiceNumber = Replace(Replace(Replace(invoiceNumber, "/", "-"), "\", "-"), ":", "-")

    ' Set PDF export path with enhanced macOS validation
    pdfPath = "/Users/narendrachowdary/development/GST(excel)/invoices(demo)/"

    ' Validate and create directory with enhanced error handling
    On Error Resume Next
    Call CreateDirectoryIfNotExists(pdfPath)
    If Err.Number <> 0 Then
        ' Try alternative path if main path fails
        pdfPath = "/Users/narendrachowdary/Desktop/"
        Call CreateDirectoryIfNotExists(pdfPath)
        If Err.Number <> 0 Then
            MsgBox "Cannot create directory for PDF export. Using Desktop as fallback.", vbExclamation, "Directory Warning"
        End If
    End If
    On Error GoTo PDFExportError

    ' Full filename with path (ensure clean filename)
    If cleanInvoiceNumber = "" Then cleanInvoiceNumber = "GST_Invoice_" & Format(Now, "yyyymmdd_hhmmss")
    fullPath = pdfPath & cleanInvoiceNumber & ".pdf"

    ' Validate the full path length (macOS has path length limits)
    If Len(fullPath) > 255 Then
        cleanInvoiceNumber = "Invoice_" & Format(Now, "yyyymmdd")
        fullPath = pdfPath & cleanInvoiceNumber & ".pdf"
    End If

    ' Delete any existing temporary sheet to avoid errors
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("DuplicateInvoiceTemp").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create a temporary duplicate of the invoice sheet (ENHANCED METHOD)
    Application.DisplayAlerts = False

    ' Copy the original sheet to create duplicate
    originalWs.Copy After:=originalWs

    ' Get reference to the newly created sheet (more reliable method)
    Set duplicateWs = Nothing
    On Error Resume Next
    Set duplicateWs = ThisWorkbook.Sheets(originalWs.Index + 1)
    On Error GoTo 0

    ' Fallback method if the above fails
    If duplicateWs Is Nothing Then
        Set duplicateWs = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    End If

    ' Ensure we have a valid duplicate sheet
    If duplicateWs Is Nothing Then
        Application.DisplayAlerts = True
        MsgBox "Failed to create duplicate sheet for PDF export.", vbCritical, "PDF Export Error"
        Exit Sub
    End If

    duplicateWs.Name = "DuplicateInvoiceTemp"
    Application.DisplayAlerts = True

    ' Change the header on the duplicate sheet to "DUPLICATE"
    duplicateWs.Range("A1").Value = "DUPLICATE"

    ' Ensure both sheets have identical content except for the header
    ' Copy all data from original to duplicate (except A1) - UPDATED RANGE TO O40
    ' Use PasteSpecial with xlPasteValues to avoid warehouse reference issues
    On Error Resume Next
    originalWs.Range("A2:O40").Copy
    duplicateWs.Range("A2").PasteSpecial Paste:=xlPasteValues
    duplicateWs.Range("A2").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo PDFExportError

    ' OPTIMIZE PDF LAYOUT - Updated for new row structure (ends at row 38)
    ' Set print area and page setup for the original sheet - OPTIMIZED FOR ENHANCED LAYOUT AND SCALING
    On Error Resume Next  ' Handle macOS PageSetup compatibility issues
    originalWs.PageSetup.PrintArea = "A1:O40"  ' Updated to include all rows up to row 40
    With originalWs.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.15)  ' Reduced margins for more content space
        .RightMargin = Application.InchesToPoints(0.15)
        .TopMargin = Application.InchesToPoints(0.15)
        .BottomMargin = Application.InchesToPoints(0.15)
        .HeaderMargin = Application.InchesToPoints(0.1)
        .FooterMargin = Application.InchesToPoints(0.1)
        .CenterHorizontally = True
        .CenterVertically = True  ' Enable vertical centering for better appearance
        .BlackAndWhite = False  ' Ensure colors are preserved
    End With
    On Error GoTo PDFExportError  ' Resume error handling

    ' Set print area and page setup for the duplicate sheet - OPTIMIZED FOR ENHANCED LAYOUT AND SCALING
    On Error Resume Next  ' Handle macOS PageSetup compatibility issues
    duplicateWs.PageSetup.PrintArea = "A1:O40"  ' Updated to include all rows up to row 40
    With duplicateWs.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.15)  ' Reduced margins for more content space
        .RightMargin = Application.InchesToPoints(0.15)
        .TopMargin = Application.InchesToPoints(0.15)
        .BottomMargin = Application.InchesToPoints(0.15)
        .HeaderMargin = Application.InchesToPoints(0.1)
        .FooterMargin = Application.InchesToPoints(0.1)
        .CenterHorizontally = True
        .CenterVertically = True  ' Enable vertical centering for better appearance
        .BlackAndWhite = False  ' Ensure colors are preserved
    End With
    On Error GoTo PDFExportError  ' Resume error handling

    ' ENHANCED PDF EXPORT with better quality and error handling
    On Error GoTo PDFExportError

    ' Apply PDF-optimized formatting before export
    On Error Resume Next
    Call OptimizeForPDFExport(originalWs)
    Call OptimizeForPDFExport(duplicateWs)
    On Error GoTo PDFExportError

    ' Verify we only have the two invoice sheets we want to export
    Dim totalSheets As Integer
    totalSheets = ThisWorkbook.Sheets.Count

    ' macOS-Compatible PDF Export Method
    On Error GoTo PDFExportError

    ' ENHANCED PDF EXPORT METHOD - Ensure only invoice sheets are exported
    On Error GoTo PDFExportError

    ' Verify both sheets exist before export
    If originalWs Is Nothing Or duplicateWs Is Nothing Then
        MsgBox "Error: Invoice sheets not found for PDF export.", vbCritical, "PDF Export Error"
        Exit Sub
    End If

    ' Method 1: Export both invoice sheets to a single PDF using explicit sheet names
    Dim sheetNames As Variant
    sheetNames = Array(originalWs.Name, duplicateWs.Name)

    ' Select only the two invoice sheets (Original and Duplicate)
    ThisWorkbook.Sheets(sheetNames).Select

    ' Export the selected sheets as a single PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                    Filename:=fullPath, _
                                    Quality:=xlQualityStandard, _
                                    IgnorePrintAreas:=False, _
                                    OpenAfterPublish:=False

    ' Restore worksheet formatting after PDF export
    On Error Resume Next
    Call RestoreWorksheetFormatting(originalWs)
    On Error GoTo PDFExportError

    ' Clean up the temporary duplicate sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    If Not duplicateWs Is Nothing Then
        duplicateWs.Delete
    End If
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Select the original invoice sheet
    originalWs.Select

    ' Success message with detailed information
    MsgBox "✅ Invoice exported successfully as a 2-page PDF!" & vbCrLf & vbCrLf & _
           "📄 Page 1: ORIGINAL (for recipient)" & vbCrLf & _
           "📄 Page 2: DUPLICATE (for driver/transport)" & vbCrLf & vbCrLf & _
           "📁 File: " & cleanInvoiceNumber & ".pdf" & vbCrLf & _
           "📂 Location: " & pdfPath, vbInformation, "PDF Export Complete"
    Exit Sub

PDFExportError:
    ' Enhanced PDF export error handling with fallback method
    If Err.Number <> 0 Then
        ' Clean up the temporary sheet first
        On Error Resume Next
        Application.DisplayAlerts = False
        If Not duplicateWs Is Nothing Then duplicateWs.Delete
        Application.DisplayAlerts = True
        On Error GoTo 0

        ' Try fallback method: Export only the original sheet
        On Error Resume Next
        Dim fallbackPath As String
        fallbackPath = Replace(fullPath, ".pdf", "_single.pdf")

        originalWs.Select
        originalWs.ExportAsFixedFormat Type:=xlTypePDF, _
                                       Filename:=fallbackPath, _
                                       Quality:=xlQualityStandard, _
                                       IgnorePrintAreas:=False, _
                                       OpenAfterPublish:=False

        If Err.Number = 0 Then
            ' Fallback succeeded
            MsgBox "PDF Export Successful (Single Page)!" & vbCrLf & _
                   "File: " & Dir(fallbackPath) & vbCrLf & _
                   "Location: " & Left(fallbackPath, InStrRev(fallbackPath, "/")) & vbCrLf & vbCrLf & _
                   "Note: Only the original invoice was exported due to macOS compatibility.", _
                   vbInformation, "PDF Export Complete"
            originalWs.Select
            Exit Sub
        End If
        On Error GoTo 0
    End If

    ' If fallback also failed, show detailed error
    Dim macOSErrorMsg As String
    macOSErrorMsg = "PDF Export Failed (macOS Troubleshooting):" & vbCrLf & vbCrLf & _
                    "Error: " & Err.Description & vbCrLf & _
                    "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
                    "macOS-Specific Solutions:" & vbCrLf & _
                    "• Check Excel permissions in System Preferences > Security & Privacy" & vbCrLf & _
                    "• Ensure the directory exists and is writable" & vbCrLf & _
                    "• Close any PDF files with the same name" & vbCrLf & _
                    "• Try exporting to Desktop first" & vbCrLf & _
                    "• Restart Excel if the issue persists"

    MsgBox macOSErrorMsg, vbCritical, "PDF Export Error"
    GoTo ErrorHandler

ErrorHandler:
    ' Enhanced error handling with detailed diagnostics
    Dim errorMsg As String
    errorMsg = "PDF Export Error Details:" & vbCrLf & vbCrLf & _
               "Error: " & Err.Description & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "PDF Path: " & pdfPath & vbCrLf & vbCrLf & _
               "Possible Solutions:" & vbCrLf & _
               "• Check if the folder path exists and is accessible" & vbCrLf & _
               "• Verify you have write permissions to the directory" & vbCrLf & _
               "• Ensure the invoice number is valid for filename" & vbCrLf & _
               "• Close any open PDF files with the same name"

    ' Ensure cleanup happens even if there's an error
    If Not duplicateWs Is Nothing Then
        Application.DisplayAlerts = False
        On Error Resume Next
        duplicateWs.Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
    End If

    ' Restore original settings
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Show detailed error message only if there was actually an error
    If Err.Number <> 0 Then
        MsgBox errorMsg, vbCritical, "PDF Export Failed"
    End If
End Sub

Public Sub OptimizeForPDFExport(ws As Worksheet)
    ' Optimize worksheet formatting specifically for PDF export - CUSTOM HORIZONTAL BORDER HANDLING
    Dim cell As Range
    On Error Resume Next

    With ws
        ' Ensure header rows 3-5 (company info) have proper borders for PDF
        ' BUT remove horizontal borders from rows 3 and 4 as requested
        .Range("A3:O5").Borders.LineStyle = xlContinuous
        .Range("A3:O5").Borders.Weight = xlMedium
        .Range("A3:O5").Borders.Color = RGB(0, 0, 0)  ' Pure black for PDF
        
        ' Remove horizontal borders from rows 3 and 4 for PDF
        With .Range("A3:O3")
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
        
        With .Range("A4:O4")
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
        
        ' Ensure data area borders are clean for PDF
        .Range("A7:O40").Borders.LineStyle = xlContinuous
        .Range("A7:O40").Borders.Weight = xlThin
        .Range("A7:O40").Borders.Color = RGB(0, 0, 0)  ' Pure black for PDF

        ' Optimize N/A display for PDF (make it less prominent)
        For Each cell In .Range("I20:N24")
            If cell.Value = "N/A" Then
                cell.Font.Color = RGB(128, 128, 128)  ' Gray instead of red for PDF
                cell.Font.Size = 8  ' Smaller font for N/A
            End If
        Next cell

        ' Optimize yellow highlighting for PDF
        For Each cell In .Range("A26:J28")  ' Amount in words section
            If cell.Interior.Color = RGB(255, 255, 0) Then  ' Yellow
                cell.Interior.Color = RGB(255, 255, 200)  ' Lighter yellow for PDF
            End If
        Next cell

        ' Ensure proper font rendering
        .Range("A1:O40").Font.Name = "Segoe UI"

        ' Set optimal row heights for PDF - UPDATED FOR COMPREHENSIVE LAYOUT OPTIMIZATION
        .Rows(2).RowHeight = 55       ' Company name header - increased for better PDF layout
        .Rows("7:10").RowHeight = 35  ' Invoice details section - increased for better PDF layout
        .Rows("17:18").RowHeight = 30 ' Two-row header structure
        .Rows("19:24").RowHeight = 38 ' Item rows - increased for better PDF layout
        .Rows(25).RowHeight = 50      ' Total quantity section - increased for better PDF layout
        .Rows("26:33").RowHeight = 32 ' Tax summary and totals - increased for better PDF layout
        .Rows("12:16").RowHeight = 35 ' Party details - increased for better PDF layout
        .Rows(34).RowHeight = 55      ' Signature headers - increased for better PDF layout
        .Rows("37:39").RowHeight = 45 ' Signature space - increased for better PDF layout
    End With

    On Error GoTo 0
End Sub

Public Sub RestoreWorksheetFormatting(ws As Worksheet)
    ' Restore worksheet formatting after PDF export - MAINTAIN CUSTOM HORIZONTAL BORDER REMOVAL
    On Error Resume Next

    With ws
        ' Ensure header rows 3-5 maintain proper borders for Excel editing
        .Range("A3:O5").Borders.LineStyle = xlContinuous
        .Range("A3:O5").Borders.Weight = xlMedium
        .Range("A3:O5").Borders.Color = RGB(0, 0, 0)
        
        ' Remove horizontal borders from rows 3 and 4 after restoration
        With .Range("A3:O3")
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
        
        With .Range("A4:O4")
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
        
        ' Restore borders for data area
        .Range("A7:O40").Borders.LineStyle = xlContinuous
        .Range("A7:O40").Borders.Weight = xlThin
        .Range("A7:O40").Borders.Color = RGB(0, 0, 0)

        ' Restore original yellow highlighting for editing
        .Range("A26").Interior.Color = RGB(255, 255, 0)  ' Terms and Conditions header
        .Range("A29").Interior.Color = RGB(255, 255, 0)  ' Terms and Conditions header
    End With

    On Error GoTo 0
End Sub
