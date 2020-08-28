Public Function CheckIfFileExists(FilePath As String)

On Error GoTo ExitWithError

If FilePath = "" Then
    CheckIfFileExists = ""
    Exit Function
End If
If Dir(FilePath) <> "" Then
    CheckIfFileExists = "Bulletin Saved and Uploaded, Ready to Save and Send to Store Comm"
Else
    CheckIfFileExists = "Bulletin not Saved, Not Ready to Send to Store Comm"
End If

Exit Function
ExitWithError:
    CheckIfFileExists = "File not accessible"
End Function
