Function GetTrackingInfo(trackingNumber As String, carrier As String) As String
    Dim htmlDoc As Object
    Dim htmlElement As Object
    Dim xmlHttp As Object
    
    ' Create an XMLHTTP object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Construct the URL based on the carrier
    Dim trackingUrl As String
    
    Select Case UCase(carrier)
        Case "USPS"
            trackingUrl = "https://tools.usps.com/go/TrackConfirmAction?qtc_tLabels1=" & trackingNumber
        Case "UPS"
            trackingUrl = "https://www.ups.com/track?tracknum=" & trackingNumber
        Case "FEDEX"
            trackingUrl = "https://www.fedex.com/fedextrack/?tracknumbers=" & trackingNumber
        Case Else
            GetTrackingInfo = "Unknown carrier"
            Exit Function
    End Select
    
    ' Open a connection to the tracking URL
    xmlHttp.Open "GET", trackingUrl, False
    xmlHttp.send
    
    ' Check if the request was successful (status code 200)
    If xmlHttp.Status = 200 Then
        ' Create an HTML document object and load the response
        Set htmlDoc = CreateObject("htmlfile")
        htmlDoc.body.innerHTML = xmlHttp.responseText
        
        ' Extract tracking information using HTML tags
        Select Case UCase(carrier)
            Case "USPS"
                Set htmlElement = htmlDoc.getElementById("tc-hits")
                If Not htmlElement Is Nothing Then
                    GetTrackingInfo = "USPS Status: " & htmlElement.innerText
                Else
                    GetTrackingInfo = "No USPS tracking information available"
                End If
            Case "UPS"
                ' Search for UPS status in the response text
                If InStr(xmlHttp.responseText, "In Transit") > 0 Then
                    GetTrackingInfo = "UPS Status: In Transit"
                ElseIf InStr(xmlHttp.responseText, "Delivered") > 0 Then
                    GetTrackingInfo = "UPS Status: Delivered"
                Else
                    GetTrackingInfo = "UPS Status: Unknown"
                End If
            Case "FEDEX"
                ' Search for FedEx status in the response text
                If InStr(xmlHttp.responseText, "In transit") > 0 Then
                    GetTrackingInfo = "FedEx Status: In Transit"
                ElseIf InStr(xmlHttp.responseText, "Delivered") > 0 Then
                    GetTrackingInfo = "FedEx Status: Delivered"
                Else
                    GetTrackingInfo = "FedEx Status: Unknown"
                End If
        End Select
    Else
        ' Handle request failure
        GetTrackingInfo = carrier & " HTTP request failed. Status code: " & xmlHttp.Status
    End If
End Function
