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
            Continue For
        End If

        ' Check if the file exists in the search folder or its subdirectories
        If fileName <> "" Then
            fileExists = Dir(searchFolder & "\" & fileName, vbNormal) <> ""
        End If

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
    Next cell

    MsgBox "File copying completed. Check the Error sheet for any missing files."
End Sub

Function IsValidFileName(ByVal fileName As String) As Boolean
    Dim invalidChars() As String
    Dim invalidChar As Variant

    ' List of characters that are not allowed in file names
    invalidChars = Split("\ / : * ? "" < > |", " ")

    ' Check if the file name contains any invalid characters
    For Each invalidChar In invalidChars
        If InStr(fileName, invalidChar) > 0 Then
            IsValidFileName = False
            Exit Function
        End If
    Next invalidChar

    IsValidFileName = True
End Function
