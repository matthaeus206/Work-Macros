Option Explicit

Public Function GetCarrierFromTracking(ByVal trackingNumber As String) As String
    Dim tn As String
    tn = UCase$(Trim$(trackingNumber))
    tn = Replace(tn, " ", "")
    tn = Replace(tn, "-", "")

    If tn = "" Then
        GetCarrierFromTracking = "Unknown"
        Exit Function
    End If

    ' UPS: 1Z + 16 characters
    If Left$(tn, 2) = "1Z" And Len(tn) = 18 Then
        GetCarrierFromTracking = "UPS"
        Exit Function
    End If

    ' USPS Express: 2 letters + 9 digits + US
    If Len(tn) = 13 _
       And Right$(tn, 2) = "US" _
       And IsNumeric(Mid$(tn, 3, 9)) _
       And Not IsNumeric(Left$(tn, 2)) Then
        GetCarrierFromTracking = "USPS"
        Exit Function
    End If

    ' USPS common 20–22 digit formats, usually starting with 9
    If IsNumeric(tn) And Len(tn) >= 20 And Len(tn) <= 22 Then
        If Left$(tn, 1) = "9" Then
            GetCarrierFromTracking = "USPS"
        Else
            GetCarrierFromTracking = "Ambiguous FedEx/USPS"
        End If
        Exit Function
    End If

    ' FedEx common numeric formats
    If IsNumeric(tn) Then
        Select Case Len(tn)
            Case 12, 15
                GetCarrierFromTracking = "FedEx"
                Exit Function
            Case 20 To 22
                GetCarrierFromTracking = "Ambiguous FedEx/USPS"
                Exit Function
        End Select
    End If

    ' DHL Express: common 10-digit format
    If Len(tn) = 10 And IsNumeric(tn) Then
        GetCarrierFromTracking = "DHL"
        Exit Function
    End If

    ' DHL eCommerce / DHL-style prefixes
    If Left$(tn, 2) = "3S" _
       Or Left$(tn, 4) = "JVGL" _
       Or Left$(tn, 3) = "JJD" Then
        GetCarrierFromTracking = "DHL"
        Exit Function
    End If

    GetCarrierFromTracking = "Unknown"
End Function
