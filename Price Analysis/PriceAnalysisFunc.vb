' === Pricing Analysis Functions ===
' --- 1. Gross Margin (%) ---
Function GrossMargin(Revenue As Double, COGS As Double) As Double
    If Revenue = 0 Then
        GrossMargin = 0
    Else
        GrossMargin = ((Revenue - COGS) / Revenue) * 100
    End If
End Function

' --- 2. Contribution Margin (Value) ---
Function ContributionMargin(SellingPrice As Double, VariableCost As Double) As Double
    ContributionMargin = SellingPrice - VariableCost
End Function

' --- 2b. Contribution Margin (%) ---
Function ContributionMarginPercent(SellingPrice As Double, VariableCost As Double) As Double
    If SellingPrice = 0 Then
        ContributionMarginPercent = 0
    Else
        ContributionMarginPercent = ((SellingPrice - VariableCost) / SellingPrice) * 100
    End If
End Function

' --- 3. Break-Even Volume (Units) ---
Function BreakEvenUnits(FixedCosts As Double, SellingPrice As Double, VariableCost As Double) As Double
    Dim ContributionPerUnit As Double
    ContributionPerUnit = SellingPrice - VariableCost
    If ContributionPerUnit = 0 Then
        BreakEvenUnits = 0
    Else
        BreakEvenUnits = FixedCosts / ContributionPerUnit
    End If
End Function

' --- 3b. Break-Even Revenue (Dollars) ---
Function BreakEvenRevenue(FixedCosts As Double, SellingPrice As Double, VariableCost As Double) As Double
    Dim CMPercent As Double
    If SellingPrice = 0 Then
        BreakEvenRevenue = 0
    Else
        CMPercent = (SellingPrice - VariableCost) / SellingPrice
        If CMPercent = 0 Then
            BreakEvenRevenue = 0
        Else
            BreakEvenRevenue = FixedCosts / CMPercent
        End If
    End If
End Function

' --- 4. Price Elasticity of Demand ---
Function PriceElasticity(ChangeInQuantityPercent As Double, ChangeInPricePercent As Double) As Double
    If ChangeInPricePercent = 0 Then
        PriceElasticity = 0
    Else
        PriceElasticity = ChangeInQuantityPercent / ChangeInPricePercent
    End If
End Function
