' Confidence Interval for a Mean (Z) with known sigma or large n
Function ConfidenceIntervalMean(mean As Double, sigma As Double, n As Double, alpha As Double) As String
    Dim z As Double: z = WorksheetFunction.NormSInv(1 - alpha / 2)
    Dim margin As Double: margin = z * (sigma / Sqr(n))
    ConfidenceIntervalMean = "[" & Round(mean - margin, 4) & ", " & Round(mean + margin, 4) & "]"
End Function

' Confidence Interval for Proportion
Function ConfidenceIntervalProportion(p As Double, n As Double, alpha As Double) As String
    Dim z As Double: z = WorksheetFunction.NormSInv(1 - alpha / 2)
    Dim margin As Double: margin = z * Sqr((p * (1 - p)) / n)
    ConfidenceIntervalProportion = "[" & Round(p - margin, 4) & ", " & Round(p + margin, 4) & "]"
End Function

' Z-test for two means (known population std dev)
Function ZTestMeans(mean1 As Double, mean2 As Double, sigma1 As Double, sigma2 As Double, n1 As Double, n2 As Double) As Double
    Dim z As Double
    z = (mean1 - mean2) / Sqr((sigma1 ^ 2 / n1) + (sigma2 ^ 2 / n2))
    ZTestMeans = z
End Function

' Z-test for two proportions
Function ZTestProportions(p1 As Double, p2 As Double, n1 As Double, n2 As Double) As Double
    Dim p_combined As Double: p_combined = (p1 * n1 + p2 * n2) / (n1 + n2)
    Dim z As Double
    z = (p1 - p2) / Sqr(p_combined * (1 - p_combined) * (1 / n1 + 1 / n2))
    ZTestProportions = z
End Function

' Compare p-value to alpha (returns "Reject H0" or "Fail to Reject H0")
Function HypothesisDecision(pValue As Double, alpha As Double) As String
    If pValue < alpha Then
        HypothesisDecision = "Reject H0"
    Else
        HypothesisDecision = "Fail to Reject H0"
    End If
End Function
