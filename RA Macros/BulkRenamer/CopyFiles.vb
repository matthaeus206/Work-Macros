Sub CopyFiles()
    Dim sourceFolder As String
    Dim destFolder As String
    Dim rng As Range
    Dim cell As Range
    Dim filename As String
    Dim i As Long
    Dim notFoundList As String
    Dim fileExtension As String
    Dim batchSize As Long
    Dim filesToCopy As Collection
    Dim file As Variant
    
    ' Get input values from user
    sourceFolder = InputBox("Enter source folder path:")
    destFolder = InputBox("Enter destination folder path:")
    fileExtension = InputBox("Enter file extension:")
    batchSize = InputBox("Enter batch size (recommended: 1000):")
    Set rng = Application.InputBox("Select cells with search terms:", Type:=8)
    
    ' Disable alerts and screen updating to prevent popups and improve performance
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' Initialize collection for files to copy
    Set filesToCopy = New Collection
    
    ' Loop through each cell in selected range
    For Each cell In rng
        ' Loop through files in source folder
        filename = Dir(sourceFolder & "\*" & cell.Value & "*" & fileExtension)
        Do While filename <> ""
            ' Add file to collection
            filesToCopy.Add sourceFolder & "\" & filename
            filename = Dir()
        Loop
    Next cell
    
    ' Loop through files in batches and copy them to destination folder
    i = 0
    For Each file In filesToCopy
        FileCopy file, destFolder & "\" & GetFileNameFromPath(file)
        i = i + 1
        
        ' Check if batch size limit is reached
        If i Mod batchSize = 0 Then
            ' Reset the counter and prevent interruption from user input
            Application.EnableCancelKey = xlDisabled
            DoEvents
            Application.EnableCancelKey = xlInterrupt
        
            ' Update the progress to the user
            Application.StatusBar = "Copying files: " & i & " files copied"
        End If
    Next file
    
    ' Enable alerts and screen updating
    Application.DisplayAlerts = True
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
    
    ' Clear the status bar
    Application.StatusBar = ""
End Sub

Function GetFileNameFromPath(ByVal fullPath As String) As String
    ' Retrieve the file name from the full file path
    GetFileNameFromPath = Mid(fullPath, InStrRev(fullPath, "\")
