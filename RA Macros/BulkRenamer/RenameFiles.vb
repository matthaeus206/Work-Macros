Sub RenameFiles()
    Dim fso As Object, files As Object, file As Object
    Dim currentName As String, newName As String
    Dim folderPath As String, findText As String, replaceText As String

    folderPath = InputBox("Enter the directory path where the files are located:")
    If folderPath = "" Then Exit Sub
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    findText = InputBox("Enter the text to find in the file names:")
    If findText = "" Then Exit Sub

    replaceText = InputBox("Enter the text to replace with in the file names:")

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set files = fso.GetFolder(folderPath).Files

    For Each file In files
        currentName = file.Name
        newName = Replace(currentName, findText, replaceText)

        If currentName <> newName Then
            If Dir(folderPath & newName) = "" Then
                On Error Resume Next
                Name folderPath & currentName As folderPath & newName
                If Err.Number <> 0 Then
                    MsgBox "Failed to rename " & currentName & ": " & Err.Description, vbExclamation
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                MsgBox "Skipping rename: " & newName & " already exists.", vbExclamation
            End If
        End If
    Next file
End Sub
