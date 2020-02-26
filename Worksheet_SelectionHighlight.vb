Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Application.CutCopyMode = False Then
Application.Calculate
End If
End Sub
