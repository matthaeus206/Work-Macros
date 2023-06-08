Sub CopyFiles()
    Dim sourceFolder As String
    Dim destFolder As String
    Dim rng As Range
    Dim cell As Range
    Dim filename As String
    Dim i As Long
    Dim notFoundList As String
    Dim fileExtension As String
    Dim filesToCopy As New Collection
    
    ' Get input values from user
    sourceFolder = InputBox("Enter source folder path:")
    destFolder = InputBox("Enter destination folder path:")
    fileExtension = InputBox("Enter file extension:")
    Set rng = Application.InputBox("Select cells with search terms:", Type:=8)
    
    ' Disable screen updating and calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Create FileSystemObject
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Loop through each cell in selected range
    For Each cell In rng
        ' Loop through files in source folder
        filename = Dir(sourceFolder & "\*" & cell.Value & "*" & fileExtension)
        Do While filename <> ""
            ' Add the file to the list of files to be copied
            filesToCopy.Add sourceFolder & "\" & filename
            filename = Dir()
        Loop
    Next cell
    
    ' Copy the files in batches
    Dim batchSize As Long
    batchSize = 100 ' Adjust the batch size as needed
    i = 0
    For i = 1 To filesToCopy.Count Step batchSize
        ' Get the current batch of files
        Dim batchFiles As New Collection
        Dim j As Long
        For j = i To i + batchSize - 1
            If j <= filesToCopy.Count Then
                batchFiles.Add filesToCopy(j)
            End If
        Next j
        
        ' Copy the batch of files to the destination folder
        For Each filename In batchFiles
            fso.CopyFile filename, destFolder & "\" & fso.GetFileName(filename)
        Next filename
    Next i
    
    ' Enable screen updating and calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Release resources
    Set fso = Nothing
    
    ' Display message with number of files copied
    MsgBox i - 1 & " file(s) copied."
End Sub
