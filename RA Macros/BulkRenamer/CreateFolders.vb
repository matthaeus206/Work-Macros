Sub CreateFoldersFromList()
    Dim myFolder As String
    Dim myCell As Range
    Dim myValue As String
    Dim myPath As String
    
    ' Get the path of the directory where the folders will be created
    myPath = Worksheets("Create Folders").Range("F2").Value
    
    ' Get the range of cells containing the folder names from the user
    On Error Resume Next ' allow the user to cancel the selection without causing an error
    Set myCell = Application.InputBox("Select the range of cells containing the folder names:", Type:=8)
    On Error GoTo 0 ' turn off error handling
    
    ' Check if the user selected a range
    If myCell Is Nothing Then
        Exit Sub
    End If
    
    ' Loop through each cell in the range and create a folder
    For Each myCell In myCell.Cells
        myValue = myCell.Value
        If myValue <> "" Then
            myFolder = myPath & "\" & Trim(myValue)
            If Len(Dir(myFolder, vbDirectory)) = 0 Then
                MkDir myFolder
            End If
        End If
    Next myCell
    
    ' Notify the user that the folders have been created
    MsgBox "Folders created successfully!"
End Sub

Sub ProcessPDFFileNames()

    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim strDirectory As String
    Dim strOldFile As String
    Dim strNewFile As String
    
    'Get the directory path from a dialog box
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        strDirectory = .SelectedItems(1)
    End With
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(strDirectory)
    
    For Each objFile In objFolder.files
        strOldFile = objFile.Name
        strNewFile = UCase(strOldFile)
        If strOldFile <> strNewFile Then
            On Error Resume Next
            objFile.Name = strNewFile
            On Error GoTo 0
        End If
    Next objFile
    
    Set objFile = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing
    
    MsgBox "File names have been capitalized successfully."
End Sub
