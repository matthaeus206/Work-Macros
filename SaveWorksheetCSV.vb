Sub CreateCSVFileWithSaveAsDialog()
    Dim ws As Worksheet
    Dim savePath As String
    Dim fileName As String

    ' Set the worksheet to the desired sheet name
    Set ws = ThisWorkbook.Sheets("UPLOAD Template FOR GRAPHICS")

    ' Prompt the user to choose a destination path and file name
    Dim chosenFile As Variant
    chosenFile = Application.GetSaveAsFilename(InitialFileName:="Untitled.csv", FileFilter:="CSV Files (*.csv), *.csv", Title:="Save As")

    ' Check if the user canceled the operation
    If chosenFile <> "False" Then
        ' Extract the chosen path and file name from the full path
        savePath = Left(chosenFile & "\", InStrRev(chosenFile, "\"))
        fileName = Mid(chosenFile, InStrRev(chosenFile, "\") + 1, InStrRev(chosenFile, ".") - InStrRev(chosenFile, "\") - 1)

        ' Save the worksheet as a CSV file
        ws.SaveAs fileName:=savePath & fileName & ".csv", FileFormat:=xlCSV, CreateBackup:=False

        ' Inform the user that the file has been saved
        MsgBox "CSV file saved successfully at: " & savePath & fileName & ".csv"
    Else
        ' Inform the user that the operation was canceled
        MsgBox "Operation canceled by the user."
    End If
End Sub
