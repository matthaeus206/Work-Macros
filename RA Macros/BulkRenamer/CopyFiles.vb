Sub CopyFiles()
    Dim sourceFolder As String
    Dim destFolder As String
    Dim fileExtension As String
    Dim userRange As Range
    Dim cell As Range
    Dim fileNamePrefix As String
    Dim errorList As Collection
    Dim errorsSheet As Worksheet
    Dim destinationPath As String
    Dim sourceFile As String
    Dim destFile As String
    Dim foundFile As Boolean

    ' Input source folder, destination folder, file extension, and user-defined range
    sourceFolder = InputBox("Enter the source folder path:")
    destFolder = InputBox("Enter the destination folder path:")
    fileExtension = InputBox("Enter the file extension (e.g., txt, xls, xlsx):")
    Set userRange = Application.InputBox("Select the user-defined range:", Type:=8)

    ' Create a collection to store error file names
    Set errorList = New Collection

    ' Loop through each cell in the user-defined range
    For Each cell In userRange
        ' Get the first 5 digits of the file name
        fileNamePrefix = Left(cell.Value, 5)

        ' Construct the full path of the source file
        sourceFile = sourceFolder & "\" & fileNamePrefix & "." & fileExtension

        ' Construct the full path of the destination file
        destFile = destFolder & "\" & fileNamePrefix & "." & fileExtension

        ' Check if the file exists in the source folder
        If Dir(sourceFile) <> "" Then
            ' File found, copy to destination
            FileCopy sourceFile, destFile
        Else
            ' File not found, add to error list
            errorList.Add fileNamePrefix
        End If
    Next cell

    ' Create a new worksheet for errors
    Set errorsSheet = Worksheets.Add
    errorsSheet.Name = "Errors Found"

    ' Paste the error file names in the new worksheet
    For i = 1 To errorList.Count
        errorsSheet.Cells(i, 1).Value = errorList(i)
    Next i

    MsgBox "Copying files completed. Check 'Errors Found' sheet for any errors.", vbInformation
End Sub
