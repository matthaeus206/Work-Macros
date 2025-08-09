Function GetCarrierFromTracking(trackingNumber As String) As String
    Dim tn As String
    tn = Trim(trackingNumber)
    tn = Replace(tn, " ", "") ' Remove spaces
    tn = Replace(tn, "-", "") ' Remove hyphens
    
    ' --- UPS ---
    ' Typical: starts with "1Z", 18 chars total
    If UCase(Left(tn, 2)) = "1Z" And Len(tn) = 18 Then
        GetCarrierFromTracking = "UPS"
        Exit Function
    End If
    
    ' --- FedEx ---
    ' Ground/Home: 12 digits
    ' Express: 15 digits
    ' Newer: 20–22 digits
    If Len(tn) = 12 And IsNumeric(tn) Then
        GetCarrierFromTracking = "FedEx"
        Exit Function
    ElseIf Len(tn) = 15 And IsNumeric(tn) Then
        GetCarrierFromTracking = "FedEx"
        Exit Function
    ElseIf (Len(tn) >= 20 And Len(tn) <= 22) And IsNumeric(tn) Then
        GetCarrierFromTracking = "FedEx"
        Exit Function
    End If
    
    ' --- USPS ---
    ' 20–22 digits, all numeric
    ' Some start with 9 and have 22 digits
    If (Len(tn) >= 20 And Len(tn) <= 22) And IsNumeric(tn) Then
        GetCarrierFromTracking = "USPS"
        Exit Function
    End If
    ' USPS Express: starts with 2 letters + 9 digits + US
    If Len(tn) = 13 And _
       Mid(UCase(tn), 11, 2) = "US" And _
       IsNumeric(Mid(tn, 3, 9)) Then
        GetCarrierFromTracking = "USPS"
        Exit Function
    End If
    
    ' --- DHL ---
    ' 10 digits all numeric
    If Len(tn) = 10 And IsNumeric(tn) Then
        GetCarrierFromTracking = "DHL"
        Exit Function
    End If
    ' DHL Express often: starts with 3S, JVGL, or JJD
    If Left(UCase(tn), 2) = "3S" Or _
       Left(UCase(tn), 4) = "JVGL" Or _
       Left(UCase(tn), 3) = "JJD" Then
        GetCarrierFromTracking = "DHL"
        Exit Function
    End If
    
    ' If no match
    GetCarrierFromTracking = "Unknown"
End Function
