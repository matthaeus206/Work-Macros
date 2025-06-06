Sub Mail_storecomm()
Application.ScreenUpdating = False
'This sends out emails to storecomm and attaches required files.
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim Ans As Long

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    ' Standard Font size and typeface
    Sheets("Ready to Deploy").Select
    Cells.Select
    With Selection.font
        .Name = "Arial"
        .Size = 16
    End With

    Set Sourcewb = ActiveWorkbook

    'Copy the ActiveSheet to a new workbook
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
    TempFileName = "Formatted for AX Promo " & Format(Now, "m-dd-yy")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
            
    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            ' Change email to correct recipient
            .to = "storecomm@bartelldrugs.com"
            '.to = "matt.walker@bartelldrugs.com"
            ' CC mark becker
            .CC = "mark.becker@bartelldrugs.com"
            .BCC = ""
            .Subject = "Weekly Add-Drop " & Format(Date, "m/dd/yyyy")
            .Body = "Hi," & vbCrLf & vbCrLf & "Could you please send this out to the stores?" & vbCrLf & vbCrLf & "If you have any questions, feel free to ask me." & vbCrLf & vbCrLf & "Thanks!"
            .Attachments.Add ("\\bdshare\buyers\Add-Drop\WEEKLY ADD DROPS\" & Format(Date, "yyyy") & "\BUY-Weekly Add-Drop " & Format(Date, "m.dd.yyyy") & ".pdf")
            .Attachments.Add ("\\bdshare\buyers\Add-Drop\WEEKLY ADD DROPS\" & Format(Date, "yyyy") & "\Weekly Add-Drop " & Format(Date, "m.dd.yyyy") & ".pdf")
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            '.Send   'or use
            .Display
        End With
        On Error GoTo 0
    '   Create Error Message box if I forget to create Buy Weekly Add Drop word doc.
        'On Error GoTo ErrMsg
        'ErrMsg:    MsgBox ("Missing Word File"),On Error GoTo 0
        .Close savechanges:=False
    End With

    'Delete the file you have send
    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = False
        .EnableEvents = True
    End With
    
    'Message Box
    Ans = MsgBox("Complete", vbOKOnly, "Complete")
    
    End Sub
