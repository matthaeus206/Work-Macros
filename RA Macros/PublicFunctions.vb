Public Function CheckIfFileExists(FilePath As String)
' This function is for checking if filepaths are valid and returns appropriate message.
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

Public Function AddPercentage(percentage, price As Long)
' This adds percentage to a price
    If price >= 0 Then
        AddPercentage = price * (1 + percentage)
   End If
End Function

Function LineCounter(Reference As Integer, previous As Integer)
'Work in progress
If Reference <> "" Then
    LineCounter = previous + 1
Else: LineCounter = previous
End If
End Function

Public Function PercentIncrease(Initial As Integer, _
                                Final As Integer)
'This Function calculates the percentage increase of 2 sets of numbers.
PercentIncrease = (Final - Initial) / Initial

End Function

Public Function PercentDecrease(OldNum As Integer, _
                                NewNum As Integer)
'This Function calculates the percentage decrease of 2 sets of numbers.
PercentDecrease = (NewNum - OldNum) / OldNum

End Function

Public Function WildCard(UPC As String)
    
If UPC = "" Then
    WildCard = ""
    ElseIf UPC <> "" Then
        WildCard = "*" & UPC & "*,"
End If

End Function

Option Explicit

Function SalesTax(TotalInvoice As Double, TaxPercentage As Double)

'Formula used for finding the sales tax
SalesTax = TotalInvoice - TotalInvoice / (1 + TaxPercentage)

End Function

Function CheckDigit(ByRef dblChars As Double, _
ByRef HowMany As Long) As String
Dim N As Long
Dim lngLen As Long
Dim lngSumR As Long
Dim lngSumL As Long
Dim lngTotal As Long
Dim strTemp As String

strTemp = CStr(dblChars)
lngLen = Len(strTemp)
'Confirm that entry is correct.
If lngLen <> HowMany Then
CheckDigit = "Incorrect Entry"
Exit Function
End If

'Add first set of numbers starting from right.
For N = lngLen To 1 Step -2
lngSumR = lngSumR + Mid(strTemp, N, 1)
Next
lngSumR = lngSumR * 3

'Add second set of numbers.
'starting 2nd character from right.
For N = (lngLen - 1) To 1 Step -2
lngSumL = lngSumL + Mid(strTemp, N, 1)
Next
N = lngSumR + lngSumL

'Round up
lngTotal = (N Mod 10)
lngTotal = 10 - lngTotal + N

CheckDigit = strTemp & (lngTotal - N)
End Function

Public Function CopyFiles(ByVal sourceRange As Range, ByVal destinationPath As String) As Boolean
' Copies items from one folder to another
    Dim sourcePaths As Variant
    Dim i As Long
    
    sourcePaths = sourceRange.Value
    
    For i = LBound(sourcePaths, 1) To UBound(sourcePaths, 1)
        On Error GoTo CopyFileError
        FileCopy sourcePaths(i, 1), destinationPath & "\" & Dir(sourcePaths(i, 1))
        On Error GoTo 0
    Next i
    
    CopyFiles = True
    Exit Function

CopyFileError:
    CopyFiles = False
End Function

Public Function InsertImage(filePathRange As Range)
' This takes images from a list of paths and inserts it in active sheet.
    Dim objShape As Object
    Dim filePath As String

    For Each cell In filePathRange
        filePath = cell.Value

        'Insert the image into the active worksheet
        Set objShape = ActiveSheet.Shapes.AddPicture(filePath, msoFalse, msoTrue, 0, 0, -1, -1)

        'Adjust the size of the image as needed
        objShape.Width = objShape.Width * 0.5
		objShape.Height = objShape.Height * 0.5
    Next cell
End Function

Function CopyFilesToFolder(filePathsRange As Range, destFolderPath As String) As Boolean
    Dim filePath As Variant
    Dim fileName As String
    Dim copySuccess As Boolean
    
    ' Check if the destination folder exists
    If Dir(destFolderPath, vbDirectory) = "" Then
        CopyFilesToFolder = False
        Exit Function
    End If
    
    ' Loop through each file path in the range and copy the file to the destination folder
    copySuccess = True
    For Each filePath In filePathsRange
        ' Get the file name from the path
        fileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
        
        ' Copy the file to the destination folder
        On Error Resume Next
        FileCopy filePath, destFolderPath & "\" & fileName
        If Err.Number <> 0 Then
            copySuccess = False
        End If
        On Error GoTo 0
    Next filePath
    
    ' If no errors occurred during the copy process, return True
    If copySuccess = True Then
        CopyFilesToFolder = True
    Else
        CopyFilesToFolder = False
    End If
End Function

Public Function RemoveLeadingDigits(ByVal inputString As String) As String
    ' This function removes any leading digits from a string
    Dim i As Long
    Dim str As String
    Dim num As Variant
    
    ' Find the position of the first non-digit character in the string
    For i = 1 To Len(inputString)
        If Not IsNumeric(Mid(inputString, i, 1)) Then
            Exit For
        End If
    Next i
    
    ' Remove the leading digits and any subsequent spaces
    str = Trim(Mid(inputString, i))
    
    ' Replace any errors with an empty string
    On Error Resume Next
    num = Application.WorksheetFunction.Substitute(inputString, Left(inputString, i - 1), "")
    If Err.Number <> 0 Then
        RemoveLeadingDigits = ""
    Else
        RemoveLeadingDigits = str
    End If
End Function

Public Sub UpdateAllQueries()
'' Speed up Macro
ThisWorkbook.Queries.FastCombine = True
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
'' Refreshes all Queries
    ActiveWorkbook.RefreshAll
'' Resets Settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
End Sub
