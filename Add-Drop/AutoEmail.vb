Sub Mail_ActiveSheet()
Application.ScreenUpdating = False
'This writes out Adds to IT and emails it.
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object

    'Copy data from formatted ax and paste in ax promo as values only.
    Sheets("Formatted for AX Promo").Select
    Range("A1:O122").Select
    Selection.Copy
    Sheets("AX Promo").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set Sourcewb = ActiveWorkbook

    'Copy the ActiveSheet to a new workbook
    Sheets("AX Promo").Copy
    Set Destwb = ActiveWorkbook

    'Determine the Excel version and file extension/format
    With Destwb
        If Val(Application.Version) < 12 Then
            'You use Excel 97-2003
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            'You use Excel 2007-2016
            Select Case Sourcewb.FileFormat
            Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If .HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56: FileExtStr = ".xls": FileFormatNum = 56
            Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
            End Select
        End If
    End With

    '    'Change all cells in the worksheet to values if you want
    '    With Destwb.Sheets(1).UsedRange
    '        .Cells.Copy
    '        .Cells.PasteSpecial xlPasteValues
    '        .Cells(1).Select
    '    End With
    '    Application.CutCopyMode = False

    'Save the new workbook/Mail it/Delete it
    TempFilePath = Environ$("temp") & "\"
    TempFileName = "AX Promo " & Format(Now, "m-dd-yy")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
            
    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            '.to = "itsupport@bartelldrugs.com"
            .to = "matt.walker@bartelldrugs.com"
            '.CC = vbNullString
            '.BCC = vbNullString
            .Subject = "PLOG Code"
            .Body = "Hi IT," & vbCrLf & vbCrLf & "When you have a chance, could you please upload these items to AX?" & vbCrLf & vbCrLf & "Thanks!"
            .Attachments.Add Destwb.FullName
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            '.Send   'or use
            .Display
        End With
        On Error GoTo 0
        .Close savechanges:=False
    End With

    'Delete the file you have send
    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
    'clear sheet
    Application.CutCopyMode = False
    Selection.ClearContents
    
End Sub
