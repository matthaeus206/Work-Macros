Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Cells.FormatConditions.Delete

    With Cells.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=ROW()=ROW(INDIRECT(""RC"",FALSE))")
        .Interior.Color = RGB(255, 255, 150) ' light yellow
    End With

    Me.Calculate
End Sub
