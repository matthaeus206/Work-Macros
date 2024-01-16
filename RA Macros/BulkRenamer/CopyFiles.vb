Sub CopyFiles()
    Dim sourceFolder As String
    Dim destFolder As String
    Dim rng As Range
    Dim Cell As Range
    Dim fileName As String
    Dim i As Long
    Dim notFoundList As String
    Dim fileExtension As String
    
    ' Get input values from user
    sourceFolder = InputBox("Enter source folder path:")
    destFolder = InputBox("Enter destination folder path:")
    fileExtension = InputBox("Enter file extension:")
    Set rng = Application.InputBox("Select cells with search terms:", Type:=8)
    
    ' Create FileSystemObject
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Disable alerts to prevent popups
    Application.DisplayAlerts = False
    
    ' Loop through each cell in selected range
    For Each Cell In rng
        ' Loop through files in source folder
        fileName = Dir(sourceFolder & "\*" & Cell.Value & "*" & fileExtension)
        If fileName = "" Then
            ' File not found, add to not found list
            notFoundList = notFoundList & Cell.Value & vbCrLf
        Else
            ' Copy file to destination folder
            fso.CopyFile sourceFolder & "\" & fileName, destFolder & "\" & fileName
            i = i + 1
        End If
        Do While fileName <> ""
            fileName = Dir()
        Loop
    Next Cell
    
    ' Enable alerts
    Application.DisplayAlerts = True
    
    ' Display message with number of files copied
    MsgBox i & " file(s) copied."
    
    ' Display list of search terms not found
    If Len(notFoundList) > 0 Then
        ' Create new sheet
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Not Found"
        
        ' Display not found list in new sheet
        ws.Range("A1").Value = "The following search terms were not found:"
        ws.Range("A2").Value = notFoundList
    End If
    
    ' Release FileSystemObject
    Set fso = Nothing
End Sub
