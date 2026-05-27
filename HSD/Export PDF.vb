Sub ExportRentalToPDF()

    Dim ws As Worksheet
    Dim pdfPath As Variant
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Rental PDF")
    
    pdfPath = Application.GetSaveAsFilename( _
                InitialFileName:="Rental.pdf", _
                FileFilter:="PDF Files (*.pdf), *.pdf")
    
    If pdfPath = False Then Exit Sub
    
    'Only print through the last row with a real Part # / Description
    lastRow = LastNonBlankDisplayRow(ws, "C", 100)
    
    'Keep header/ship-to area if there are no line items
    If lastRow < 10 Then lastRow = 10
    
    With ws.PageSetup
        .PrintArea = "A1:I" & lastRow
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

End Sub


Private Function LastNonBlankDisplayRow(ws As Worksheet, keyCol As String, maxRow As Long) As Long

    Dim r As Long
    
    For r = maxRow To 1 Step -1
        If Trim(CStr(ws.Cells(r, keyCol).value)) <> "" Then
            LastNonBlankDisplayRow = r
            Exit Function
        End If
    Next r
    
    LastNonBlankDisplayRow = 1

End Function

