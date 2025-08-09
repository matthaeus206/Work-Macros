Option Explicit

'========================
' CONFIG: Paste your keys here
'========================
Private Const USPS_USERID As String = "YOUR_USPS_USERID"

' UPS: Either paste a Bearer token, or implement the OAuth helper below and call it to fetch one.
Private Const UPS_BEARER_TOKEN As String = ""  ' e.g., "eyJhbGciOi..."  (leave empty if not using yet)
'Private Const UPS_CLIENT_ID As String = ""     ' if you want to implement token retrieval
'Private Const UPS_CLIENT_SECRET As String = ""

' FedEx bearer token (or implement OAuth fetch)
Private Const FEDEX_BEARER_TOKEN As String = ""

' DHL key/token
Private Const DHL_BEARER_TOKEN As String = ""

'========================
' Public API
'========================
Public Function GetTrackingStatus(ByVal rawTracking As String) As String
    Dim tn As String, carrier As String
    tn = NormalizeTracking(rawTracking)
    If Len(tn) = 0 Then
        GetTrackingStatus = "Invalid/empty tracking number"
        Exit Function
    End If

    carrier = GetCarrierFromTracking(tn)

    Select Case carrier
        Case "USPS"
            GetTrackingStatus = "USPS: " & USPS_Status(tn)
        Case "UPS"
            GetTrackingStatus = "UPS: " & UPS_Status(tn)
        Case "FedEx"
            GetTrackingStatus = "FedEx: " & FedEx_Status(tn)
        Case "DHL"
            GetTrackingStatus = "DHL: " & DHL_Status(tn)
        Case Else
            GetTrackingStatus = "Unknown carrier"
    End Select
End Function

'========================
' Carrier detection
'========================
Private Function GetCarrierFromTracking(ByVal tn As String) As String
    ' UPS: "1Z" + 16 more (total 18)
    If Len(tn) = 18 And UCase$(Left$(tn, 2)) = "1Z" Then
        GetCarrierFromTracking = "UPS": Exit Function
    End If

    ' USPS: 20–22 digits numeric, or 13-char (2 letters + 9 digits + "US")
    If (Len(tn) >= 20 And Len(tn) <= 22 And IsNumeric(tn)) _
       Or (Len(tn) = 13 And Mid$(UCase$(tn), 11, 2) = "US" And IsNumeric(Mid$(tn, 3, 9))) Then
        GetCarrierFromTracking = "USPS": Exit Function
    End If

    ' FedEx: 12, 15, or 20–22 digits (numeric)
    If (Len(tn) = 12 Or Len(tn) = 15 Or (Len(tn) >= 20 And Len(tn) <= 22)) And IsNumeric(tn) Then
        GetCarrierFromTracking = "FedEx": Exit Function
    End If

    ' DHL: 10 digits; or prefixes 3S / JVGL / JJD
    If (Len(tn) = 10 And IsNumeric(tn)) _
       Or Left$(UCase$(tn), 2) = "3S" _
       Or Left$(UCase$(tn), 4) = "JVGL" _
       Or Left$(UCase$(tn), 3) = "JJD" Then
        GetCarrierFromTracking = "DHL": Exit Function
    End If

    GetCarrierFromTracking = "Unknown"
End Function

Private Function NormalizeTracking(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace$(t, " ", "")
    t = Replace$(t, "-", "")
    NormalizeTracking = t
End Function

'========================
' USPS (works with USERID only)
'========================
Private Function USPS_Status(ByVal tn As String) As String
    If USPS_USERID = "" Then
        USPS_Status = "Set USPS_USERID in config."
        Exit Function
    End If

    ' Build simple TrackV2 request
    Dim url As String
    url = "https://secure.shippingapis.com/ShippingAPI.dll?API=TrackV2&XML=" & _
          URLEncode("<TrackRequest USERID='" & USPS_USERID & "'><TrackID ID='" & tn & "'/></TrackRequest>")

    Dim resp As String
    resp = HttpGet(url, "MSXML2.ServerXMLHTTP.6.0")

    If Len(resp) = 0 Then
        USPS_Status = "No response"
        Exit Function
    End If

    ' Parse TrackSummary (latest status)
    Dim summary As String, eventCity As String, eventState As String, eventTime As String
    summary = ExtractTag(resp, "TrackSummary")
    eventCity = ExtractTag(resp, "EventCity")
    eventState = ExtractTag(resp, "EventState")
    eventTime = ExtractTag(resp, "EventTime")

    Dim pieces As String
    pieces = summary
    If Len(eventCity) > 0 Or Len(eventState) > 0 Then
        pieces = pieces & IIf(Len(pieces) > 0, " — ", "") & eventCity & IIf(Len(eventState) > 0, ", " & eventState, "")
    End If
    If Len(eventTime) > 0 Then
        pieces = pieces & IIf(Len(pieces) > 0, " — ", "") & eventTime
    End If

    If Len(pieces) = 0 Then pieces = "Status not found"
    USPS_Status = pieces
End Function

'========================
' UPS (requires bearer token)
'========================
Private Function UPS_Status(ByVal tn As String) As String
    If UPS_BEARER_TOKEN = "" Then
        UPS_Status = "Provide UPS bearer token in config."
        Exit Function
    End If

    ' NOTE: Fill endpoint per your UPS developer dashboard (Tracking API).
    ' Example pattern (placeholder): https://onlinetools.ups.com/api/track/v1/details/{tn}
    Dim url As String
    url = "https://onlinetools.ups.com/api/track/v1/details/" & tn  ' <-- adjust if your account uses different base

    Dim headers As Object: Set headers = CreateObject("Scripting.Dictionary")
    headers.Add "Authorization", "Bearer " & UPS_BEARER_TOKEN
    headers.Add "Accept", "application/json"

    Dim json As String
    json = HttpGetWithHeaders(url, headers, "WinHttp.WinHttpRequest.5.1")

    If Len(json) = 0 Then
        UPS_Status = "No response"
        Exit Function
    End If

    ' Very simple extraction—adjust to your UPS JSON shape
    Dim statusStr As String
    statusStr = JsonPeek(json, """description"":") ' crude grep; replace with JSON parser if you have one
    If Len(statusStr) = 0 Then statusStr = JsonPeek(json, """status"":")
    If Len(statusStr) = 0 Then statusStr = "Status not found"
    UPS_Status = statusStr
End Function

'========================
' FedEx (requires bearer token)
'========================
Private Function FedEx_Status(ByVal tn As String) As String
    If FEDEX_BEARER_TOKEN = "" Then
        FedEx_Status = "Provide FedEx bearer token in config."
        Exit Function
    End If

    ' Fill endpoint per FedEx Tracking API (JSON).
    Dim url As String
    url = "https://apis.fedex.com/track/v1/trackingnumbers" ' <-- typical POST endpoint; some accounts differ

    Dim body As String
    body = "{""trackingInfo"":[{""trackingNumberInfo"":{""trackingNumber"":""" & tn & """}}],""includeDetailedScans"":false}"

    Dim headers As Object: Set headers = CreateObject("Scripting.Dictionary")
    headers.Add "Authorization", "Bearer " & FEDEX_BEARER_TOKEN
    headers.Add "Content-Type", "application/json"
    headers.Add "Accept", "application/json"

    Dim json As String
    json = HttpPost(url, body, headers, "WinHttp.WinHttpRequest.5.1")

    If Len(json) = 0 Then
        FedEx_Status = "No response"
        Exit Function
    End If

    Dim statusStr As String
    statusStr = JsonPeek(json, """statusDescription"":")
    If Len(statusStr) = 0 Then statusStr = JsonPeek(json, """derivedStatus"":")
    If Len(statusStr) = 0 Then statusStr = "Status not found"
    FedEx_Status = statusStr
End Function

'========================
' DHL (requires key/bearer)
'========================
Private Function DHL_Status(ByVal tn As String) As String
    If DHL_BEARER_TOKEN = "" Then
        DHL_Status = "Provide DHL token/key in config."
        Exit Function
    End If

    ' Fill endpoint per DHL tracking API.
    Dim url As String
    url = "https://api-eu.dhl.com/track/shipments?trackingNumber=" & URLEncode(tn) ' example pattern

    Dim headers As Object: Set headers = CreateObject("Scripting.Dictionary")
    headers.Add "Authorization", "Bearer " & DHL_BEARER_TOKEN
    headers.Add "Accept", "application/json"

    Dim json As String
    json = HttpGetWithHeaders(url, headers, "WinHttp.WinHttpRequest.5.1")

    If Len(json) = 0 Then
        DHL_Status = "No response"
        Exit Function
    End If

    Dim statusStr As String
    statusStr = JsonPeek(json, """status"":")
    If Len(statusStr) = 0 Then statusStr = "Status not found"
    DHL_Status = statusStr
End Function

'========================
' HTTP Helpers
'========================
Private Function HttpGet(ByVal url As String, ByVal progId As String) As String
    On Error GoTo EH
    Dim http As Object
    Set http = CreateObject(progId)
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "ExcelVBA"
    http.send
    If http.Status = 200 Then
        HttpGet = CStr(http.responseText)
    Else
        HttpGet = ""
    End If
    Exit Function
EH:
    HttpGet = ""
End Function

Private Function HttpGetWithHeaders(ByVal url As String, ByVal headers As Object, ByVal progId As String) As String
    On Error GoTo EH
    Dim http As Object, k As Variant
    Set http = CreateObject(progId)
    http.Open "GET", url, False
    For Each k In headers.Keys
        http.setRequestHeader CStr(k), CStr(headers(k))
    Next k
    http.send
    If http.Status = 200 Then
        HttpGetWithHeaders = CStr(http.responseText)
    Else
        HttpGetWithHeaders = ""
    End If
    Exit Function
EH:
    HttpGetWithHeaders = ""
End Function

Private Function HttpPost(ByVal url As String, ByVal body As String, ByVal headers As Object, ByVal progId As String) As String
    On Error GoTo EH
    Dim http As Object, k As Variant
    Set http = CreateObject(progId)
    http.Open "POST", url, False
    For Each k In headers.Keys
        http.setRequestHeader CStr(k), CStr(headers(k))
    Next k
    http.send body
    If http.Status = 200 Or http.Status = 202 Then
        HttpPost = CStr(http.responseText)
    Else
        HttpPost = ""
    End If
    Exit Function
EH:
    HttpPost = ""
End Function

'========================
' Tiny parsers (good enough for status strings)
'========================
Private Function ExtractTag(ByVal xmlText As String, ByVal tagName As String) As String
    ' crude XML slice: <tag>...</tag>
    Dim openTag As String, closeTag As String, a As Long, b As Long
    openTag = "<" & tagName & ">"
    closeTag = "</" & tagName & ">"
    a = InStr(1, xmlText, openTag, vbTextCompare)
    b = InStr(1, xmlText, closeTag, vbTextCompare)
    If a > 0 And b > a Then
        ExtractTag = HtmlDecode(Mid$(xmlText, a + Len(openTag), b - (a + Len(openTag))))
    Else
        ExtractTag = ""
    End If
End Function

Private Function HtmlDecode(ByVal s As String) As String
    Dim t As String
    t = Replace$(s, "&amp;", "&")
    t = Replace$(t, "&lt;", "<")
    t = Replace$(t, "&gt;", ">")
    t = Replace$(t, "&quot;", """")
    t = Replace$(t, "&apos;", "'")
    HtmlDecode = t
End Function

Private Function JsonPeek(ByVal json As String, ByVal keyWithQuotesAndColon As String) As String
    ' super-simple grep: finds the substring after key:
    ' Example: key = """statusDescription"":"
    Dim p As Long, q As Long, r As Long
    p = InStr(1, json, keyWithQuotesAndColon, vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(keyWithQuotesAndColon)
    ' skip spaces
    Do While p <= Len(json) And Mid$(json, p, 1) Like " " Or Mid$(json, p, 1) = Chr(9)
        p = p + 1
    Loop
    ' if next is quote, read quoted string
    If Mid$(json, p, 1) = """" Then
        p = p + 1
        q = InStr(p, json, """")
        If q > p Then
            JsonPeek = Mid$(json, p, q - p)
        End If
    Else
        ' read until comma/brace
        q = p
        Do While q <= Len(json)
            If Mid$(json, q, 1) Like "[,}]" Then Exit Do
            q = q + 1
        Loop
        JsonPeek = Trim$(Mid$(json, p, q - p))
    End If
End Function

'========================
' URL encoding for USPS XML payload
'========================
Private Function URLEncode(ByVal s As String) As String
    Dim i As Long, ch As Integer, t As String
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        Select Case ch
            Case 48 To 57, 65 To 90, 97 To 122  ' 0-9 A-Z a-z
                t = t & Chr$(ch)
            Case Else
                t = t & "%" & Right$("0" & Hex$(ch), 2)
        End Select
    Next i
    URLEncode = t
End Function
