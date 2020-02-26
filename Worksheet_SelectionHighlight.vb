Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    Application.ScreenUpdating = False
    ' Clear the color of all the cells
    Cells.Interior.ColorIndex = 0
    With Target
        ' Highlight the entire row and column that contain the active cell
        .EntireRow.Interior.ColorIndex = 8
        .EntireColumn.Interior.ColorIndex = 8
    End With
    Application.ScreenUpdating = True
End Sub
