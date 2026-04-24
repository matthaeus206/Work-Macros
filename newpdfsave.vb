Sub ExportRentalToPDF()

    Dim ws As Worksheet
    Dim pdfPath As Variant
    
    ' Set worksheet
    Set ws = ThisWorkbook.Worksheets("Rental PDF")
    
    ' Prompt user for save location
    pdfPath = Application.GetSaveAsFilename( _
                InitialFileName:="Rental.pdf", _
                FileFilter:="PDF Files (*.pdf), *.pdf")
    
    ' Exit if user cancels
    If pdfPath = False Then Exit Sub
    
    ' Page setup
    With ws.PageSetup
        .PrintArea = "A1:I35"
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    ' Export to PDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

End Sub