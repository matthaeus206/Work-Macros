Sub savepdf()
Application.ScreenUpdating = False
'Main routine to print out PDFs and send mail to storecomm and run everything.
    Dim strFname As String
    Dim strPath As String
    Dim strPathSP As String
    Dim oDoc As Worksheet
    Set oDoc = Sheets("Ready to Deploy")
strFname = "Weekly Add-Drop " & _
            Format(Date, "m.dd.yyyy")
'define the folder location to save the document
strPath = "\\bdshare\buyers\Add-Drop\WEEKLY ADD DROPS\" & Format(Date, "yyyy") & "\" & _
      strFname & ".pdf"

Call Mail_ActiveSheet
'Call FnOpeneWordDoc

    'sFile = Application.DefaultFilePath & "\" & _
      'ActiveWorkbook.Name & ".pdf"

    'Sheets("Sheet1").Select
    'ActiveSheet.PageSetup
    
'ActiveWorkbook.SaveAs Filename:=strPath & "\" &
      'strFname & ".xlsx"
    With Sheets("Ready to Deploy").PageSetup
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
      
    Sheets("Ready to Deploy").ExportAsFixedFormat Type:=xlTypePDF, _
      Filename:=strPath, Quality:=xlQualityStandard, _
      IncludeDocProperties:=True, IgnorePrintAreas:=False, _
      OpenAfterPublish:=False
Sheets("Ready to Deploy").Copy
 With Sheets("Ready to Deploy").UsedRange
 .Copy
 .PasteSpecial xlValues
 .PasteSpecial xlFormats
 End With
 Application.CutCopyMode = False
 ActiveWorkbook.SaveAs "\\bdshare\buyers\Add-Drop\WEEKLY ADD DROPS\" & Format(Date, "yyyy") & "\" & strFname & ".xlsx"
 ActiveWorkbook.SaveAs Filename:= _
        "https://bartelldrugs.sharepoint.com/sites/bartellnet/buying/Shared%20Documents/Add%20Drop/" & Format(Date, "yyyy") & "/" & strFname & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

 Sheets("Ready to Deploy").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "https://bartelldrugs.sharepoint.com/sites/bartellnet/buying/Shared%20Documents/Add%20Drop/" & Format(Date, "yyyy") & "/" & strFname & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
Call Mail_storecomm
End Sub
