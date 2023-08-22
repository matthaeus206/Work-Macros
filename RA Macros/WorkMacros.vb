'Capitalize all characters for files names in a folder - cell is referenced
Sub Capitalize()
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set Folder = objFSO.GetFolder(Range("F4").Value)


For Each File In Folder.Files
    sNewFile = File.Name
    sNewFile = UCase(File.Name)
    If (sNewFile <> File.Name) Then
        File.Move (File.ParentFolder + "\" + sNewFile)
    End If

Next
End Sub

'Rename all files in folder referencing cells 
Sub Rename()
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set Folder = objFSO.GetFolder(Range("F4").Value)


For Each File In Folder.Files
    sNewFile = File.Name
    sNewFile = Replace(sNewFile, (Range("G4").Value), (Range("H4").Value))
    If (sNewFile <> File.Name) Then
        File.Move (File.ParentFolder + "\" + sNewFile)
    End If

Next
End Sub

'Refresh PivotTables
Sub RefreshAllPivotTables()
Dim PT As PivotTable
For Each PT In ActiveSheet.PivotTables
PT.RefreshTable
Next PT
End Sub

'This code would highlight alternate rows in the selection
Sub HighlightAlternateRows()
Dim Myrange As Range
Dim Myrow As Range
Set Myrange = Selection
For Each Myrow In Myrange.Rows
   If Myrow.Row Mod 2 = 1 Then
      Myrow.Interior.Color = vbCyan
   End If
Next Myrow
End Sub

'Add prefix to each cell in selection
Sub AddPrefix()
Dim c As Range
Dim prefixValue As Variant

'Display inputbox to collect prefix text
prefixValue = Application.InputBox(Prompt:="Enter prefix:", _
    Title:="Prefix", Type:=2)

'The User clicked Cancel
If prefixValue = False Then Exit Sub

For Each c In Selection

    'Add prefix where cell is not a formula or blank
    If Not c.HasFormula And c.Value <> "" Then

        c.Value = prefixValue & c.Value

    End If

Next

End Sub

'Add suffix to each cell in selection
Sub AddSuffix()
Dim c As Range
Dim suffixValue As Variant

'Display inputbox to collect prefix text
suffixValue = Application.InputBox(Prompt:="Enter Suffix:", _
    Title:="Suffix", Type:=2)

'The User clicked Cancel
If suffixValue = False Then Exit Sub

    'Loop through each cellin selection
    For Each c In Selection

        'Add Suffix where cell is not a formula or blank
        If Not c.HasFormula And c.Value <> "" Then

            c.Value = c.Value & suffixValue

        End If

Next

End Sub

Sub CalendarExport()
'
' Calendar Export Macro
'

'
    Sheets("Sheet1").Select
    Range("G2").Select
    Selection.Copy
    Sheets("Book-test (2)").Select
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs Filename:= _
        "L:\CatMrc\ITC\APOLLO\Space Management\Matt\Calendar Sharing\Main Calendar.csv" _
        , FileFormat:=xlCSV, CreateBackup:=False
End Sub
                        
'Capitalize all characters for files names in a folder - cell is referenced
Sub Capitalize()
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set Folder = objFSO.GetFolder(Range("F4").Value)


For Each File In Folder.Files
    sNewFile = File.Name
    sNewFile = UCase(File.Name)
    If (sNewFile <> File.Name) Then
        File.Move (File.ParentFolder + "\" + sNewFile)
    End If

Next
End Sub

'Rename all files in folder referencing cells 
Sub Rename()
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set Folder = objFSO.GetFolder(Range("F4").Value)


For Each File In Folder.Files
    sNewFile = File.Name
    sNewFile = Replace(sNewFile, (Range("G4").Value), (Range("H4").Value))
    If (sNewFile <> File.Name) Then
        File.Move (File.ParentFolder + "\" + sNewFile)
    End If

Next
End Sub

'Refresh PivotTables
Sub RefreshAllPivotTables()
Dim PT As PivotTable
For Each PT In ActiveSheet.PivotTables
PT.RefreshTable
Next PT
End Sub

'This code would highlight alternate rows in the selection
Sub HighlightAlternateRows()
Dim Myrange As Range
Dim Myrow As Range
Set Myrange = Selection
For Each Myrow In Myrange.Rows
   If Myrow.Row Mod 2 = 1 Then
      Myrow.Interior.Color = vbCyan
   End If
Next Myrow
End Sub

'Add prefix to each cell in selection
Sub AddPrefix()
Dim c As Range
Dim prefixValue As Variant

'Display inputbox to collect prefix text
prefixValue = Application.InputBox(Prompt:="Enter prefix:", _
    Title:="Prefix", Type:=2)

'The User clicked Cancel
If prefixValue = False Then Exit Sub

For Each c In Selection

    'Add prefix where cell is not a formula or blank
    If Not c.HasFormula And c.Value <> "" Then

        c.Value = prefixValue & c.Value

    End If

Next

End Sub

'Add suffix to each cell in selection
Sub AddSuffix()
Dim c As Range
Dim suffixValue As Variant

'Display inputbox to collect prefix text
suffixValue = Application.InputBox(Prompt:="Enter Suffix:", _
    Title:="Suffix", Type:=2)

'The User clicked Cancel
If suffixValue = False Then Exit Sub

    'Loop through each cellin selection
    For Each c In Selection

        'Add Suffix where cell is not a formula or blank
        If Not c.HasFormula And c.Value <> "" Then

            c.Value = c.Value & suffixValue

        End If

Next

End Sub

Sub CalendarExport()
'
' Calendar Export Macro
'

'
    Sheets("Sheet1").Select
    Range("G2").Select
    Selection.Copy
    Sheets("Book-test (2)").Select
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs Filename:= _
        "L:\CatMrc\ITC\APOLLO\Space Management\Matt\Calendar Sharing\Main Calendar.csv" _
        , FileFormat:=xlCSV, CreateBackup:=False
End Sub

Sub UnhideSheets()
	Dim Sht As Worksheet
	For Each Sht In ActiveWorkbook.Worksheets
		Sht.Visible = xlSheetVisible
	Next Sht
End Sub
											
'This macro refreshes all connections in the active workbook.
Sub RefreshAllConnections()



Dim wkb As Workbook
Dim cn As WorkbookConnection

Set wkb = ThisWorkbook

For Each cn In wkb.Connections
    cn.Refresh
Next cn

End Sub
