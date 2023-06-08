Sub UpperCase()
    Dim objFSO As Object
    Dim folder As Object
    Dim file As Object
    Dim folderPath As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Prompt the user to enter the directory path using an input box
    folderPath = InputBox("Enter the directory path where the files are located:")
    
    'Exit the subroutine if the user clicks the cancel button on the input box
    If folderPath = "" Then Exit Sub
    
    'Get the folder object from the selected path
    Set folder = objFSO.GetFolder(folderPath)

    For Each file In folder.files
        sNewFile = UCase(file.Name)
        If sNewFile <> file.Name Then
            On Error Resume Next ' Ignore errors and continue loop
            file.Move (file.ParentFolder & "\" & sNewFile)
            On Error GoTo 0 ' Reset error handling
        End If
    Next
End Sub

