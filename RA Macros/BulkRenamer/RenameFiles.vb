Sub RenameFiles()
    Dim files As Object
    Dim file As Object
    Dim currentName As String
    Dim newName As String
    
    'Prompt the user to enter the directory path using an input box
    Dim folderPath As String
    folderPath = InputBox("Enter the directory path where the files are located:")
    
    'Exit the subroutine if the user clicks the cancel button on the input box
    If folderPath = "" Then Exit Sub

    'Prompt the user to enter the text to find in the file names using an input box
    Dim findText As String
    findText = InputBox("Enter the text to find in the file names:")
    
    'Exit the subroutine if the user clicks the cancel button on the input box
    If findText = "" Then Exit Sub
    
    'Get all files in the directory
    Set files = CreateObject("Scripting.FileSystemObject").GetFolder(folderPath).Files
    
    'Loop through each file and rename it
    For Each file In files
        currentName = file.Name
        
        'Check if the findText exists in the currentName
        If InStr(1, currentName, findText, vbTextCompare) > 0 Then
            'Replace findText with an empty string
            newName = Replace(currentName, findText, "")
            
            'Rename the file
            Name folderPath & "\" & currentName As folderPath & "\" & newName
        End If
    Next file
End Sub
