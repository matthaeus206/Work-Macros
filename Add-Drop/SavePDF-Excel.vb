Sub savepdf()
    Dim strFname As String
    Dim strPath As String
    Dim strPathSP As String
    Dim oDoc As Worksheet
    Set oDoc = Sheets("Ready to Deploy")
    
strFname = "Weekly Add-Drop " & _
            Format(Date, "m.dd.yyyy")
            
'define the folder location to save the document
strPath = "\\bdshare\buyers\Add-Drop\WEEKLY ADD DROPS\2020\" & _
      strFname & ".pdf"

Call Mail_ActiveSheet
      
    'sFile = Application.DefaultFilePath & "\" & _
      'ActiveWorkbook.Name & ".pdf"

    'Sheets("Sheet1").Select
    'ActiveSheet.PageSetup
    
'ActiveWorkbook.SaveAs Filename:=strPath & "\" &
      'strFname & ".xlsx"
    With Sheets("Ready to Deploy").PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
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
 ActiveWorkbook.SaveAs "\\bdshare\buyers\Add-Drop\WEEKLY ADD DROPS\2020\" & strFname & ".xlsx"
 ActiveWorkbook.SaveAs Filename:= _
        "https://bartelldrugs.sharepoint.com/sites/bartellnet/buying/Shared%20Documents/Add%20Drop/2020/" & strFname & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

 Sheets("Ready to Deploy").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "https://bartelldrugs.sharepoint.com/sites/bartellnet/buying/Shared%20Documents/Add%20Drop/2020/" & strFname & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
      
Call FnOpeneWordDoc
      
End Sub
