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
