Sub MacroSaveAsPDF()
'macro saves pdf either in the same folder where active doc is or in documents folder if file is not yet saved
'
    Dim strPath As String
    Dim strPDFname As String
 
    strPDFname = InputBox("Enter name for PDF", "File Name", "example")
    If strPDFname = "" Then 'user deleted text from inputbox, add default name
        strPDFname = "example"
    End If
    strPath = ActiveDocument.Path
    If strPath = "" Then    'doc is not saved yet
        strPath = Options.DefaultFilePath(wdDocumentsPath) & Application.PathSeparator
    Else
        'just add \ at the end
        strPath = strPath & Application.PathSeparator
    End If
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
                                       strPath & strPDFname & ".pdf", _
                                       ExportFormat:=wdExportFormatPDF, _
                                       OpenAfterExport:=False, _
                                       OptimizeFor:=wdExportOptimizeForPrint, _
                                       Range:=wdExportAllDocument, _
                                       IncludeDocProps:=True, _
                                       CreateBookmarks:=wdExportCreateWordBookmarks, _
                                       BitmapMissingFonts:=True
End Sub
