Function IdentifyCarrier(trackingNumber As String) As String

    Dim carrier As String
    Dim trackingNumberLength As Integer

    trackingNumberLength = Len(trackingNumber)

    If trackingNumberLength >= 12 And trackingNumberLength <= 14 Then
        carrier = "FedEx"
    ElseIf trackingNumberLength >= 20 And trackingNumberLength <= 22 Then
        carrier = "USPS"
    ElseIf InStr(trackingNumber, "1Z") Then
        carrier = "UPS"
    Else
        carrier = "Unknown"
    End If

    IdentifyCarrier = carrier

End Function
