Sub DelChartFormat()
'
' DelChartFormat Macro
'

'
    Range("A1").Select
    Workbooks.OpenText Filename:="C:\Users\mrcmrw\Desktop\Del.txt", Origin:= _
        1250, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=True, _
        Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), Array(2, 1), Array( _
        3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10 _
        , 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), _
        Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array( _
        23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), _
        Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array( _
        36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), _
        Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), Array( _
        49, 1), Array(50, 1), Array(51, 1), Array(52, 1), Array(53, 1), Array(54, 1), Array(55, 1), _
        Array(56, 1), Array(57, 1), Array(58, 1), Array(59, 1), Array(60, 1), Array(61, 1), Array( _
        62, 1), Array(63, 1), Array(64, 1), Array(65, 1), Array(66, 1), Array(67, 1), Array(68, 1), _
        Array(69, 1), Array(70, 1), Array(71, 1), Array(72, 1), Array(73, 1), Array(74, 1), Array( _
        75, 1), Array(76, 1), Array(77, 1), Array(78, 1), Array(79, 1), Array(80, 1), Array(81, 1), _
        Array(82, 1), Array(83, 1), Array(84, 1)), TrailingMinusNumbers:=True
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H8").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("G:K").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A3").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Rows("3:3").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Rows("3:3").Select
    With Selection
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A3").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("F:F").ColumnWidth = 15
    Columns("G:G").ColumnWidth = 14
    Columns("H:H").ColumnWidth = 18.5
    Columns("I:I").ColumnWidth = 16
    Columns("J:J").ColumnWidth = 18
    Columns("K:K").ColumnWidth = 15
    Columns("L:L").ColumnWidth = 11
    Columns("M:M").ColumnWidth = 11
    Range("A:A,E:E,F:F,G:G,H:H,I:I,J:J,K:K").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "Y / N / R"
    Range("G7").Select
    Selection.FormulaR1C1 = _
        "Delete:" & Chr(10) & Chr(10) & "Yes= Chainwide" & Chr(10) & "No= Sectional" & Chr(10) & "Relo= Relocation" & Chr(10)
    Range("H7").Select
    Selection.FormulaR1C1 = _
        "In-home Markdown:" & Chr(10) & Chr(10) & "*No In-Home Markdown" & Chr(10) & "*Funded" & Chr(10) & "*Non-Funded" & Chr(10) & "*RELO" & Chr(10)
    Range("I7").Select
    Selection.FormulaR1C1 = _
        "Recall Timing:" & Chr(10) & Chr(10) & "*Instore Date" & Chr(10) & "*Post Clearance (Markdown must =75%)" & Chr(10) & "*RELO" & Chr(10)
    Range("J7").Select
    Selection.FormulaR1C1 = _
        "Recall type:" & Chr(10) & Chr(10) & "*Funded (Non POI)" & Chr(10) & "*Pay On Inventory (POI)" & Chr(10) & "*Salvaged" & Chr(10) & "*N/A-NO recall PPC (PermPriceChange Zero)" & Chr(10) & "*RELO"
    Range("K7").Select
    ActiveCell.FormulaR1C1 = "POI Chainwide Deletes (Y):" & Chr(10) & Chr(10) & "*Push DC inventory to Stores " & Chr(10) & " *NoPush of DC Inventory to Stores" & Chr(10)
    Range("L7").Select
    ActiveCell.FormulaR1C1 = "Stores Keeping" & Chr(10) & "Item In Set"
    Range("M7").Select
    Selection.FormulaR1C1 = "Store Deletes"
    Range("A7:M7").Select
    With Selection
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("G7:K7").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("G7:K7").Select
    With Selection
        .VerticalAlignment = xlTop
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("L7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("A:A").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Range("A3:M7").Select
    Selection.Font.Bold = True
    Range("N3").Select
    Selection.End(xlToRight).Select
    ActiveCell.Offset(4, 0).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(RC[2]:RC[202],"""",R5C[2]:R5C[202])"
    Selection.Copy
    Range("L8").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("E8").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 7).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(RC[1]:RC[202],""=X"",R5C[1]:R5C[202])"
    Selection.Copy
    Range("M8").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("E8").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("L3").Select
    ActiveCell.FormulaR1C1 = " "
    Range("M3").Select
    ActiveCell.FormulaR1C1 = " "
    
    
''    Rows("8:8").Select
''    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
''    Range("E8").Select
''    ActiveCell.Offset(0, 1).Range("A1").Select
''    ActiveCell.FormulaR1C1 = "stop"
''    ActiveCell.Offset(0, 1).Range("A1").Select
''    ActiveCell.FormulaR1C1 = "stop"
''    ActiveCell.Offset(0, 1).Range("A1").Select
''    ActiveCell.FormulaR1C1 = "stop"
''    ActiveCell.Offset(0, 1).Range("A1").Select
''    ActiveCell.FormulaR1C1 = "stop"
''    ActiveCell.Offset(0, 1).Range("A1").Select
''    ActiveCell.FormulaR1C1 = "stop"
''    Range("E8").Select
''    Selection.End(xlDown).Select
''    ActiveCell.Offset(0, 1).Range("A1").Select
''    Range(Selection, Selection.End(xlUp)).Select
''    With Selection.Validation
''        .Delete
''        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
''        xlBetween, Formula1:="=$K$1:$N$1"
''        .IgnoreBlank = True
''        .InCellDropdown = True
''        .InputTitle = ""
''        .ErrorTitle = ""
''        .InputMessage = ""
''        .ErrorMessage = ""
''        .ShowInput = True
''        .ShowError = True
''    End With
''    Range("E8").Select
''    Selection.End(xlDown).Select
''    ActiveCell.Offset(0, 2).Range("A1").Select
''    Range(Selection, Selection.End(xlUp)).Select
''    With Selection.Validation
''        .Delete
''        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
''        xlBetween, Formula1:="=$O$1:$S$1"
''        .IgnoreBlank = True
''        .InCellDropdown = True
''        .InputTitle = ""
''        .ErrorTitle = ""
''        .InputMessage = ""
''        .ErrorMessage = ""
''        .ShowInput = True
''        .ShowError = True
''    End With
''    Range("E8").Select
''    Selection.End(xlDown).Select
''    ActiveCell.Offset(0, 3).Range("A1").Select
''    Range(Selection, Selection.End(xlUp)).Select
''    With Selection.Validation
''        .Delete
''        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
''        xlBetween, Formula1:="=$T$1:$W$1"
''        .IgnoreBlank = True
''        .InCellDropdown = True
''        .InputTitle = ""
''        .ErrorTitle = ""
''        .InputMessage = ""
''        .ErrorMessage = ""
''        .ShowInput = True
''        .ShowError = True
''    End With

    ActiveWindow.View = xlPageLayoutView
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = "Rite Aid Delete Chart"
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    Range("C3").Select
    ActiveWindow.SmallScroll Down:=21
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = "Rite Aid Delete Chart"
        .RightHeader = ""
        .LeftFooter = _
        "X=ITEM IS BEING DELETED FROM THIS SECTION" & Chr(10) & "N/A=DOES NOT PERTAIN TO THIS SIZE SECTION" & Chr(10) & "BLANK=ITEM REMAINS IN THIS SIZE SECTION"
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.2)
        .RightMargin = Application.InchesToPoints(0.2)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLegal
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 75
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    Application.PrintCommunication = True
    ActiveWindow.View = xlNormalView
    Range("K1:Y1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("A1").Select
    End With
End Sub
