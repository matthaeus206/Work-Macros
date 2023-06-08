Sub CopyFiles()
    Dim sourceFolder As String
    Dim destFolder As String
    Dim rng As Range
    Dim cell As Range
    Dim filename As String
    Dim i As Long
    Dim notFoundList As String
    Dim fileExtension As String
    
    ' Get input values from user
    sourceFolder = InputBox("Enter source folder path:")
    destFolder = InputBox("Enter destination folder path:")
    fileExtension = InputBox("Enter file extension:")
    Set rng = Application.InputBox("Select cells with search terms:", Type:=8)
    
    ' Disable alerts to prevent popups
    Application.DisplayAlerts = False
    
    ' Disable screen updating
    Application.ScreenUpdating = False
    
    ' Loop through each cell in selected range
    For Each cell In rng
        ' Construct the full file path
        filename = sourceFolder & "\" & cell.Value & fileExtension
        
        ' Check if the file exists
        If Dir(filename) <> "" Then
            ' Copy file to destination folder
            FileCopy filename, destFolder & "\" & cell.Value & fileExtension
            i = i + 1
        Else
            ' File not found, add to not found list
            notFoundList = notFoundList & cell.Value & vbCrLf
        End If
    Next cell
    
    ' Enable alerts
    Application.DisplayAlerts = True
    
    ' Enable screen updating
    Application.ScreenUpdating = True
    
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
    
End Sub
