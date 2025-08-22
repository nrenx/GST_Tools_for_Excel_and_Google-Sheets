Option Explicit
' ===============================================================================
' MODULE: PDFUtilities
' DESCRIPTION: Helper utilities for PDF export functionality including directory 
'              creation and macOS compatibility functions
' ===============================================================================

Public Sub CreateDirectoryIfNotExists(directoryPath As String)
    ' Robust directory creation that works across different operating systems
    ' Handles both Windows and macOS compatibility issues
    Dim fso As Object
    On Error GoTo DirectoryError

    ' Try FileSystemObject first (works on most systems)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(directoryPath) Then
        fso.CreateFolder directoryPath
        ' Directory created successfully
    Else
        ' Directory already exists
    End If
    Set fso = Nothing
    Exit Sub

DirectoryError:
    ' Fallback method for macOS or when FileSystemObject fails
    On Error Resume Next
    Set fso = Nothing

    ' Try using MkDir as fallback (more compatible with macOS)
    If Dir(directoryPath, vbDirectory) = "" Then
        MkDir directoryPath
        If Err.Number <> 0 Then
            ' Don't throw error - let the PDF export attempt to continue
        End If
    End If

    On Error GoTo 0
End Sub

Private Function GetMacOSCompatiblePDFPath() As String
    ' Get a reliable PDF export path for macOS
    Dim testPath As String

    ' Try the intended directory first
    testPath = "/Users/narendrachowdary/development/GST(excel)/invoices(demo)/"
    If Dir(testPath, vbDirectory) <> "" Then
        GetMacOSCompatiblePDFPath = testPath
        Exit Function
    End If

    ' Fallback to Desktop
    testPath = "/Users/narendrachowdary/Desktop/"
    If Dir(testPath, vbDirectory) <> "" Then
        GetMacOSCompatiblePDFPath = testPath
        Exit Function
    End If

    ' Last resort - Documents folder
    testPath = "/Users/narendrachowdary/Documents/"
    GetMacOSCompatiblePDFPath = testPath
End Function

Public Sub SimplePDFExportForMacOS()
    ' Simplified, highly reliable PDF export for macOS
    Dim ws As Worksheet
    Dim invoiceNumber As String
    Dim pdfPath As String
    Dim fullPath As String
    On Error GoTo SimpleExportError

    Set ws = ThisWorkbook.Worksheets("GST_Tax_Invoice_for_interstate")
    invoiceNumber = Trim(ws.Range("C7").Value)

    If invoiceNumber = "" Then
        MsgBox "Please enter an invoice number before exporting to PDF.", vbExclamation, "Missing Invoice Number"
        Exit Sub
    End If

    ' Use Desktop as the most reliable path on macOS
    pdfPath = "/Users/narendrachowdary/Desktop/"
    fullPath = pdfPath & Replace(invoiceNumber, "/", "-") & ".pdf"

    ' Simple, single-sheet export (most reliable on macOS)
    ws.Select
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                           Filename:=fullPath, _
                           Quality:=xlQualityStandard, _
                           IgnorePrintAreas:=False, _
                           OpenAfterPublish:=False

    MsgBox "PDF exported successfully to Desktop!" & vbCrLf & _
           "File: " & Replace(invoiceNumber, "/", "-") & ".pdf", _
           vbInformation, "PDF Export Complete"
    Exit Sub

SimpleExportError:
    MsgBox "Simple PDF export failed: " & Err.Description & vbCrLf & _
           "Please check file permissions and try again.", vbCritical, "Export Error"
End Sub
