Function HoltLinearForecast()

    Dim alpha As Double
    Dim beta As Double
    Dim userRange As Range
    Dim cell As Range
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Prompt for alpha
    alpha = Application.InputBox( _
        Prompt:="Enter α (alpha), the smoothing factor for LEVEL (0 to 1)", _
        Title:="Holt Linear Forecast", Type:=1)

    ' Prompt for beta
    beta = Application.InputBox( _
        Prompt:="Enter β (beta), the smoothing factor for TREND (0 to 1)", _
        Title:="Holt Linear Forecast", Type:=1)

    ' Prompt for data range
    Set userRange = Application.InputBox( _
        Prompt:="Select the range of actual demand values (must be a single column)", _
        Title:="Holt Linear Forecast", Type:=8)

    Dim level As Double, trend As Double
    Dim i As Long
    Dim rCount As Long
    rCount = userRange.Rows.Count

    level = userRange.Cells(1, 1).Value
    trend = userRange.Cells(2, 1).Value - userRange.Cells(1, 1).Value

    ' Output header
    userRange.Cells(1, 1).Offset(0, 1).Value = "Holt Forecast"

    ' Initial forecast
    userRange.Cells(2, 1).Offset(0, 1).Value = level + trend

    For i = 3 To rCount
        Dim actual As Double
        actual = userRange.Cells(i, 1).Value

        Dim prevLevel As Double
        prevLevel = level

        level = alpha * actual + (1 - alpha) * (level + trend)
        trend = beta * (level - prevLevel) + (1 - beta) * trend

        userRange.Cells(i, 1).Offset(0, 1).Value = level + trend
    Next i

    MsgBox "Forecast complete. Output written to column next to input.", vbInformation

End Function

Function HoltWintersAdditiveForecast()

    Dim alpha As Double
    Dim beta As Double
    Dim gamma As Double
    Dim seasonLength As Long
    Dim userRange As Range
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Get smoothing parameters
    alpha = Application.InputBox( _
        Prompt:="Enter α (alpha), smoothing for LEVEL (0 to 1)", _
        Title:="Holt-Winters Forecast", Type:=1)

    beta = Application.InputBox( _
        Prompt:="Enter β (beta), smoothing for TREND (0 to 1)", _
        Title:="Holt-Winters Forecast", Type:=1)

    gamma = Application.InputBox( _
        Prompt:="Enter γ (gamma), smoothing for SEASONAL component (0 to 1)", _
        Title:="Holt-Winters Forecast", Type:=1)

    seasonLength = Application.InputBox( _
        Prompt:="Enter season length (e.g., 12 for monthly, 4 for weekly with quarterly pattern)", _
        Title:="Holt-Winters Forecast", Type:=1)

    ' Get the data range
    Set userRange = Application.InputBox( _
        Prompt:="Select the range of actual demand values (single column)", _
        Title:="Holt-Winters Forecast", Type:=8)

    Dim rCount As Long
    rCount = userRange.Rows.Count

    If rCount <= seasonLength Then
        MsgBox "Range must be longer than the season length!", vbCritical
        Exit Function
    End If

    Dim level() As Double, trend() As Double, season() As Double
    Dim t As Long
    ReDim level(1 To rCount)
    ReDim trend(1 To rCount)
    ReDim season(1 To seasonLength)

    ' Initialize
    Dim avgSeason As Double
    avgSeason = Application.WorksheetFunction.Average(userRange.Cells(1, 1).Resize(seasonLength))
    For t = 1 To seasonLength
        season(t) = userRange.Cells(t, 1).Value - avgSeason
    Next t

    level(seasonLength) = avgSeason
    trend(seasonLength) = (userRange.Cells(seasonLength + 1, 1).Value - userRange.Cells(1, 1).Value) / seasonLength

    ' Output header
    userRange.Cells(1, 1).Offset(0, 1).Value = "HW Forecast"

    For t = seasonLength + 1 To rCount
        Dim actual As Double: actual = userRange.Cells(t, 1).Value
        Dim sIndex As Long: sIndex = ((t - 1) Mod seasonLength) + 1

        level(t) = alpha * (actual - season(sIndex)) + (1 - alpha) * (level(t - 1) + trend(t - 1))
        trend(t) = beta * (level(t) - level(t - 1)) + (1 - beta) * trend(t - 1)
        season(sIndex) = gamma * (actual - level(t)) + (1 - gamma) * season(sIndex)

        userRange.Cells(t, 1).Offset(0, 1).Value = level(t) + trend(t) + season(sIndex)
    Next t

    MsgBox "Holt-Winters forecast complete. Results written to adjacent column.", vbInformation

End Function

Function CrostonsForecast(demandRange As Range, alpha As Double) As Double
    Dim n As Long: n = demandRange.Count
    Dim forecast As Double: forecast = 0
    Dim intervals As Double: intervals = 0
    Dim lastDemand As Double: lastDemand = 0
    Dim intervalCount As Long: intervalCount = 1
    
    Dim i As Long
    For i = 1 To n
        Dim demand As Double: demand = demandRange.Cells(i, 1).Value
        If demand > 0 Then
            forecast = forecast + alpha * (demand - forecast)
            intervals = intervals + alpha * (intervalCount - intervals)
            intervalCount = 1
        Else
            intervalCount = intervalCount + 1
        End If
    Next i
    
    If intervals = 0 Then
        CrostonsForecast = 0
    Else
        CrostonsForecast = forecast / intervals
    End If
End Function

Function IsOutlier(value As Double, mean As Double, stdDev As Double) As Boolean
    If Abs((value - mean) / stdDev) > 3 Then
        IsOutlier = True
    Else
        IsOutlier = False
    End If
End Function

Function RollingMAPE(actualRange As Range, forecastRange As Range, period As Long) As Double
    Dim totalError As Double, i As Long
    Dim startRow As Long: startRow = actualRange.Count - period + 1
    For i = startRow To actualRange.Count
        If actualRange.Cells(i, 1).Value <> 0 Then
            totalError = totalError + Abs((actualRange.Cells(i, 1).Value - forecastRange.Cells(i, 1).Value) / actualRange.Cells(i, 1).Value)
        End If
    Next i
    RollingMAPE = (totalError / period) * 100
End Function
