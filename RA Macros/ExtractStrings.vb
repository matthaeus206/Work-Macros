Sub ExtractStrings()
    Dim directoryPath As String
    Dim fileFormat As String
    Dim extractSheet As Worksheet
    Dim fileName As String
    Dim fileContent As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim extractedString As String
    
    ' Prompt user for directory path and file format
    directoryPath = InputBox("Enter the directory path:")
    fileFormat = InputBox("Enter the file format (e.g., jpg, png):")
    
    ' Create a new worksheet named "extract"
    Set extractSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    extractSheet.Name = "extract"
    
    ' Start searching for files
    SearchFiles directoryPath, fileFormat, extractSheet
End Sub

Sub SearchFiles(ByVal folderPath As String, ByVal fileFormat As String, ByVal extractSheet As Worksheet)
    Dim fileName As String
    Dim fileContent As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim extractedString As String
    
    ' Search files in the current folder
    fileName = Dir(folderPath & "\*." & fileFormat)
    Do While fileName <> ""
        ' Check if the file starts with "L:\\" and ends with the specified format
        If Left(fileName, 3) = "L:\\" And (UCase(Right(fileName, Len(fileFormat))) = UCase(fileFormat)) Then
            ' Read the file content
            Open folderPath & "\" & fileName For Input As #1
            fileContent = Input$(LOF(1), #1)
            Close #1
            
            ' Find the position of the desired string
            startPos = InStr(1, fileContent, "YourStartStringHere")
            endPos = InStr(1, fileContent, "YourEndStringHere")
            
            ' Extract the string
            If startPos > 0 And endPos > startPos Then
                extractedString = Mid(fileContent, startPos, endPos - startPos)
                
                ' Insert the extracted string into the "extract" worksheet
                extractSheet.Cells(extractSheet.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = extractedString
            End If
        End If
        
        ' Move to the next file
        fileName = Dir
    Loop
    
    ' Search files in subfolders
    Dim subfolderPath As String
    subfolderPath = Dir(folderPath & "\*", vbDirectory)
    Do While subfolderPath <> ""
        If subfolderPath <> "." And subfolderPath <> ".." Then
            SearchFiles folderPath & "\" & subfolderPath, fileFormat, extractSheet
        End If
        subfolderPath = Dir
    Loop
End Sub
