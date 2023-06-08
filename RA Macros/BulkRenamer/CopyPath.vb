Sub CopyPaths()

    'Declare variables
    Dim destDir As String
    Dim filePath As String
    Dim rangeSelection As Range
    Dim currentCell As Range
    
    'Get user input
    destDir = InputBox("Enter the destination directory:")
    
    'Select range of file paths
    On Error Resume Next
    Set rangeSelection = Application.InputBox(prompt:="Select the range of file paths:", Type:=8)
    On Error GoTo 0
    
    'Loop through selected cells
    If Not rangeSelection Is Nothing Then
        For Each currentCell In rangeSelection.Cells
            'Get file path from cell
            filePath = currentCell.Value
            
            'Copy file (ignore errors if file not found)
            On Error Resume Next
            fileCopy filePath, destDir & "\" & Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
            On Error GoTo 0
        Next currentCell
        
        'Display message when done
        MsgBox "All files copied."
    End If

End Sub
