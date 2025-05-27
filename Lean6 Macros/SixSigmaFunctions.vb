' DPU = Total Defects / Total Units
Function DPU(totalDefects As Double, totalUnits As Double) As Double
    If totalUnits = 0 Then
        DPU = 0
    Else
        DPU = totalDefects / totalUnits
    End If
End Function

' DPMO = (Total Defects / (Units * Opportunities)) * 1,000,000
Function DPMO(totalDefects As Double, totalUnits As Double, oppsPerUnit As Double) As Double
    If totalUnits = 0 Or oppsPerUnit = 0 Then
        DPMO = 0
    Else
        DPMO = (totalDefects / (totalUnits * oppsPerUnit)) * 1000000
    End If
End Function

' Yield = Good Units / Total Units
Function YieldRate(goodUnits As Double, totalUnits As Double) As Double
    If totalUnits = 0 Then
        YieldRate = 0
    Else
        YieldRate = goodUnits / totalUnits
    End If
End Function

' RTY = Product of yields (pass an array of yields as input)
Function RTY(ParamArray yields() As Variant) As Double
    Dim i As Integer
    Dim result As Double: result = 1
    For i = LBound(yields) To UBound(yields)
        result = result * yields(i)
    Next i
    RTY = result
End Function

' Sigma Level from DPMO
Function SigmaLevel(dpmoVal As Double) As Double
    If dpmoVal <= 0 Or dpmoVal >= 1000000 Then
        SigmaLevel = 0
    Else
        SigmaLevel = WorksheetFunction.NormSInv(1 - dpmoVal / 1000000) + 1.5
    End If
End Function

' Cp = (USL - LSL) / (6 * sigma)
Function Cp(usl As Double, lsl As Double, sigma As Double) As Double
    If sigma = 0 Then
        Cp = 0
    Else
        Cp = (usl - lsl) / (6 * sigma)
    End If
End Function

' Cpk = Min((USL - mean) / 3sigma, (mean - LSL) / 3sigma)
Function Cpk(usl As Double, lsl As Double, mean As Double, sigma As Double) As Double
    If sigma = 0 Then
        Cpk = 0
    Else
        Cpk = Application.Min((usl - mean) / (3 * sigma), (mean - lsl) / (3 * sigma))
    End If
End Function

' Standard Deviation (population)
Function StdDevPopulation(dataRange As Range) As Double
    StdDevPopulation = WorksheetFunction.StDev_P(dataRange)
End Function

' Range = Max - Min
Function RangeCalc(dataRange As Range) As Double
    RangeCalc = WorksheetFunction.Max(dataRange) - WorksheetFunction.Min(dataRange)
End Function

' Takt Time = Available Time / Customer Demand
Function TaktTime(availableTime As Double, customerDemand As Double) As Double
    If customerDemand = 0 Then
        TaktTime = 0
    Else
        TaktTime = availableTime / customerDemand
    End If
End Function
