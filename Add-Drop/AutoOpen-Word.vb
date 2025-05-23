Public Sub AutoOpen()

Dim oDoc As Document
Dim strFname As String
Dim strPath As String
Dim strPathSP As String
Set oDoc = ActiveDocument
'define the format of the filename - with the year and time as a prefix
'adding seconds to the time will ensure that each time you use the macro you will get
'a new copy.
strFname = "BUY-Weekly Add-Drop " & _
            Format(Date, "m.dd.yyyy")

'define the folder location to save the document
strPath = "\\bdshare\buyers\Add-Drop\WEEKLY ADD DROPS\2020\"
With oDoc
    'print the document to the current printer
    '.PrintOut
    'save the document in Word docx format
    .SaveAs strPath & strFname & ".docx", FileFormat:=wdFormatDocumentDefault
    'save the document as PDF format in the same folder
    .SaveAs strPath & strFname & ".pdf", FileFormat:=wdFormatPDF
    'Save to sharepoint.
    .SaveAs "https://bartelldrugs.sharepoint.com/sites/bartellnet/buying/Shared%20Documents/Add%20Drop/" & Format(Date, "yyyy") & "/" & strFname & ".pdf", FileFormat:=wdFormatPDF
End With

End Sub
