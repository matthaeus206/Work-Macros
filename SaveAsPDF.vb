Sub savepdf()

Dim oDoc As Document
Dim strFname As String
Dim strPath As String
Set oDoc = ActiveDocument
'define the format of the filename - with the year and time as a prefix
'adding seconds to the time will ensure that each time you use the macro you will get
'a new copy.
strFname = "BUY-Weekly Add-Drop " & _
            Format(Date, "mm.dd.yyyy")

'define the folder location to save the document
strPath = "C:\Users\matt.walker\Desktop\Test\"
With oDoc
    'print the document to the current printer
    '.PrintOut
    'save the document in Word docx format
    .SaveAs strPath & strFname & ".docx", FileFormat:=wdFormatDocumentDefault
    'save the document as PDF format in the same folder
    .SaveAs strPath & strFname & ".pdf", FileFormat:=wdFormatPDF
    
End With
End Sub
