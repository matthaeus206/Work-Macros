Option Explicit

Private Function Code128Glyph(ByVal value As Long) As String
    If value < 0 Or value > 106 Then
        Err.Raise vbObjectError + 514, , "Invalid Code 128 value"
    End If

    If value <= 94 Then
        Code128Glyph = ChrW(value + 32)
    Else
        Code128Glyph = ChrW(value + 100)
    End If
End Function

Public Function Code128B(ByVal inputText As String) As String
    Dim i As Long
    Dim checksum As Long
    Dim charValue As Long
    Dim encoded As String

    ' Start Code B = 104
    checksum = 104
    encoded = ChrW(204)

    For i = 1 To Len(inputText)
        charValue = AscW(Mid$(inputText, i, 1)) - 32

        If charValue < 0 Or charValue > 95 Then
            Err.Raise vbObjectError + 513, , "Invalid character for Code 128 B"
        End If

        checksum = checksum + (charValue * i)

        ' Code Set B data characters
        encoded = encoded & Code128Glyph(charValue)
    Next i

    checksum = checksum Mod 103

    ' Checksum character
    encoded = encoded & Code128Glyph(checksum)

    ' Stop character = 106
    encoded = encoded & ChrW(206)

    Code128B = encoded
End Function
