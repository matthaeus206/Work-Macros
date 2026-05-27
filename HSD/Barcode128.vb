Option Explicit

Private Function Code128Glyph(ByVal value As Long) As String
    If value < 0 Or value > 106 Then
        Err.Raise vbObjectError + 514, , "Invalid Code 128 value: " & value
    End If

    Select Case value
        Case 0
            Code128Glyph = ChrW(194)   ' Libre Barcode 128: avoid literal space
        Case 1 To 94
            Code128Glyph = ChrW(value + 32)
        Case 95 To 106
            Code128Glyph = ChrW(value + 100)
    End Select
End Function

Public Function Code128B(ByVal inputText As String) As String
    Dim i As Long
    Dim checksum As Long
    Dim charValue As Long
    Dim encoded As String
    Dim currentChar As String

    If Len(inputText) = 0 Then
        Err.Raise vbObjectError + 512, , "Input text cannot be blank"
    End If

    ' Start Code B = 104
    checksum = 104
    encoded = Code128Glyph(104)

    For i = 1 To Len(inputText)
        currentChar = Mid$(inputText, i, 1)
        charValue = AscW(currentChar) - 32

        If charValue < 0 Or charValue > 95 Then
            Err.Raise vbObjectError + 513, , _
                "Invalid character for Code 128 B: " & currentChar
        End If

        checksum = checksum + (charValue * i)
        encoded = encoded & Code128Glyph(charValue)
    Next i

    checksum = checksum Mod 103

    encoded = encoded & Code128Glyph(checksum)
    encoded = encoded & Code128Glyph(106)

    Code128B = encoded
End Function
