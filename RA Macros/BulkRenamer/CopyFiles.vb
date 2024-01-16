Sub CopyFiles()
    Dim sourceFolder As String
    Dim destFolder As String
    Dim rng As Range
    Dim cell As Range
    Dim filename As String
    Dim i As Long
    Dim notFoundList As String
    Dim fileExtension As String
    
    ' Get input values from user
    sourceFolder = InputBox("Enter source folder path:")
    destFolder = InputBox("Enter destination folder path:")
    fileExtension = InputBox("Enter file extension:")
    
    ' Validate folder paths
    If Not ValidateFolder(sourceFolder, "Source") Or Not ValidateFolder(destFolder, "Destination") Then
        Exit Sub
    End If
    
    Set rng = Application.InputBox("Select cells with search terms:", Type:=8)
    
    ' Create FileSystemObject
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Disable alerts to prevent popups
    Application.DisplayAlerts = False
    
    ' Loop through each cell in selected range
    For Each cell In rng
        ' Loop through files in source folder
        filename = Dir(fso.BuildPath(sourceFolder, "*" & cell.Value & "*" & fileExtension))
        If filename = "" Then
            ' File not found, add to not found list
            notFoundList = notFoundList & cell.Value & vbCrLf
        Else
            ' Copy file to destination folder
            On Error Resume Next
            fso.CopyFile fso.BuildPath(sourceFolder, filename), fso.BuildPath(destFolder, filename)
            On Error GoTo 0
            i = i + 1
        End If
        Do While filename <> ""
            filename = Dir()
        Loop
    Next cell
    
    ' Enable alerts
    Application.DisplayAlerts = True
    
    ' Display message with number of files copied
    MsgBox i & " file(s) copied."
    
    ' Display list of search terms not found
    If Len(notFoundList) > 0 Then
        ' Create new sheet
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Not Found"
        
        ' Display not found list in new sheet
        ws.Range("A1").Value = "The following search terms were not found:"
        ws.Range("A2").Value = notFoundList
    End If
    
    ' Release FileSystemObject
    Set fso = Nothing
End Sub

Function ValidateFolder(folderPath As String, folderType As String) As Boolean
    ' Validate if the folder exists
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        MsgBox folderType & " folder does not exist.", vbExclamation
        ValidateFolder = False
    Else
        ValidateFolder = True
    End If
    
    Set fso = Nothing
End Function
