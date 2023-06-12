Sub CopyMatchingFiles()
    Dim searchRange As Range
    Dim searchFolder As String
    Dim destinationFolder As String
    Dim errorSheet As Worksheet
    Dim errorCell As Range
    Dim fileName As String
    Dim fileExists As Boolean
    Dim cell As Range
    
    ' Set the user-defined search range
    Set searchRange = Application.InputBox("Enter the range containing file names:", Type:=8)
    
    ' Set the user-defined search folder
    searchFolder = Application.InputBox("Enter the search folder:", Type:=2)
    
    ' Set the user-defined destination folder
    destinationFolder = Application.InputBox("Enter the destination folder:", Type:=2)
    
    ' Create a new error sheet or clear existing one
    On Error Resume Next
    Set errorSheet = ThisWorkbook.Sheets("Error")
    On Error GoTo 0
    
    If errorSheet Is Nothing Then
        Set errorSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        errorSheet.Name = "Error"
    Else
        errorSheet.Cells.Clear
    End If
    
    ' Loop through each cell in the search range
    For Each cell In searchRange
        fileName = Trim(cell.Value)
        fileExists = False
        
        ' Check if the file name is valid
        If Not IsValidFileName(fileName) Then
            ' If the file name is invalid, write it in the error sheet
            Set errorCell = errorSheet.Cells(errorSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
            errorCell.Value = fileName
            GoTo NextIteration
        End If
        
        ' Search for the file in the search folder and its subfolders
        fileExists = SearchFile(searchFolder, fileName)
        
        ' If the file exists, copy it to the destination folder
        If fileExists Then
            On Error Resume Next
            FileCopy searchFolder & "\" & fileName, destinationFolder & "\" & fileName
            On Error GoTo 0
        Else
            ' If the file doesn't exist, write the filename in the error sheet
            Set errorCell = errorSheet.Cells(errorSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
            errorCell.Value = fileName
        End If
        
NextIteration:
    Next cell
    
    MsgBox "File copying completed. Check the Error sheet for any missing files."
End Sub

Function SearchFile(ByVal folderPath As String, ByVal fileName As String) As Boolean
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim found As Boolean
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' Check the files in the current folder
    For Each file In folder.Files
        If UCase(file.Name) = UCase(fileName) Then
            found = True
            Exit For
        End If
    Next file
    
    ' Recursively search through subfolders
    If Not found Then
        For Each subFolder In folder.Subfolders
            found = SearchFile(subFolder.Path, fileName)
            If found Then Exit For
        Next subFolder
    End If
    
    SearchFile = found
End Function

Function IsValidFileName(ByVal fileName As String) As Boolean
    Dim invalidChars()
