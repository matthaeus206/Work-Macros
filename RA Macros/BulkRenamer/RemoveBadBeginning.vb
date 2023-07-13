Sub RemoveTextFromFilenames()
    Dim folderPath As String
    Dim searchText As String
    Dim file As String
    Dim newName As String
    
    ' Prompt the user to enter the folder path
    folderPath = InputBox("Enter the folder path:")
    
    ' Prompt the user to enter the search text
    searchText = InputBox("Enter the text to remove:")
    
    ' Validate if the folder path is provided
    If folderPath = "" Then
        MsgBox "Folder path is not specified.", vbCritical
        Exit Sub
    End If
    
    ' Validate if the search text is provided
    If searchText = "" Then
        MsgBox "Search text is not specified.", vbCritical
        Exit Sub
    End If
    
    ' Check if the specified folder path exists
    If Not Dir(folderPath, vbDirectory) = vbNullString Then
        ' Loop through each file in the folder
        file = Dir(folderPath & "\*.*")
        Do While file <> ""
            ' Check if the search text is present in the filename
            If InStr(1, file, searchText, vbTextCompare) > 0 Then
                ' Remove the text and everything before it
                newName = Mid(file, InStr(file, searchText) + Len(searchText))
                
                ' Rename the file
                Name folderPath & "\" & file As folderPath & "\" & newName
            End If
            file = Dir
        Loop
        MsgBox "File renaming completed.", vbInformation
    Else
        MsgBox "Invalid folder path.", vbCritical
    End If
End Sub
