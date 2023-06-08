Sub CopyFiles()
    Dim sourceFolder As String
    Dim destFolder As String
    Dim rng As Range
    Dim cell As Range
    Dim fileExtension As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim i As Long
    Dim notFoundList As String

    ' Disable screen updating and set calculation to manual
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Get input values from user
    sourceFolder = InputBox("Enter source folder path:")
    destFolder = InputBox("Enter destination folder path:")
    fileExtension = InputBox("Enter file extension:")
    Set rng = Application.InputBox("Select cells with search terms:", Type:=8)
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Disable alerts to prevent popups
    Application.DisplayAlerts = False
    
    ' Loop through each cell in selected range
    For Each cell In rng
        ' Loop through files in source folder
        Set folder = fso.GetFolder(sourceFolder)
        
        For Each file In folder.Files
            If fso.GetExtensionName(file.Name) = fileExtension And file.Name Like "*" & cell.Value & "*" Then
                ' Copy file to destination folder
                file.Copy destFolder & "\" & file.Name
                i = i + 1
            End If
        Next file
        
        Set folder = Nothing
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
    
    ' Enable screen updating and reset calculation mode
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
