Sub savepdf()
    Dim strFname As String
    Dim strPath As String
    Dim strPathSP As String
    Dim oDoc As Worksheet
    Set oDoc = ActiveSheet
    
strFname = "Weekly Add-Drop " & _
            Format(Date, "m.dd.yyyy")
            
'define the folder location to save the document
strPath = "\\bdshare\buyers\Add-Drop\WEEKLY ADD DROPS\2020\" & _
      strFname & ".pdf"
      
    'sFile = Application.DefaultFilePath & "\" & _
      'ActiveWorkbook.Name & ".pdf"

    'Sheets("Sheet1").Select
    'ActiveSheet.PageSetup
    
'ActiveWorkbook.SaveAs Filename:=strPath & "\" &
      'strFname & ".xlsx"
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
      
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
      Filename:=strPath, Quality:=xlQualityStandard, _
      IncludeDocProperties:=True, IgnorePrintAreas:=False, _
      OpenAfterPublish:=True

ActiveSheet.Copy
 With ActiveSheet.UsedRange
 .Copy
 .PasteSpecial xlValues
 .PasteSpecial xlFormats
 End With
 Application.CutCopyMode = False
 ActiveWorkbook.SaveAs "\\bdshare\buyers\Add-Drop\WEEKLY ADD DROPS\2020\" & strFname & ".xlsx"
 ActiveWorkbook.SaveAs "//bartelldrugs.sharepoint.com/:f:/r/sites/bartellnet/buying/Shared%20Documents/Add%20Drop/2020?csf=1&e=CoqESL/" & strFname & ".xlsx"

End Sub
