' === Pricing Models ===

' --- 1. Cost-Plus Pricing ---
' Returns: Selling Price based on cost and markup %
Function CostPlusPrice(Cost As Double, MarkupPercent As Double) As Double
    CostPlusPrice = Cost * (1 + MarkupPercent / 100)
End Function

' --- 2. Value-Based Pricing ---
' Returns: Suggested Price based on perceived value and desired value capture %
Function ValueBasedPrice(PerceivedValue As Double, CaptureRatePercent As Double) As Double
    ValueBasedPrice = PerceivedValue * (CaptureRatePercent / 100)
End Function

' --- 3. Competitive Pricing ---
' Returns: Suggested Price based on competitor pricing and position %
' Example: Undercut = -5, Match = 0, Premium = +10
Function CompetitivePrice(CompetitorPrice As Double, AdjustmentPercent As Double) As Double
    CompetitivePrice = CompetitorPrice * (1 + AdjustmentPercent / 100)
End Function

' --- 4. Dynamic Pricing ---
' Returns: Adjusted price based on demand factor (e.g., 1.2 for high demand, 0.8 for low)
Function DynamicPrice(BasePrice As Double, DemandFactor As Double) As Double
    DynamicPrice = BasePrice * DemandFactor
End Function

' --- 5. Markdown Optimization ---
' Returns: New price after markdown
Function MarkdownPrice(OriginalPrice As Double, MarkdownPercent As Double) As Double
    MarkdownPrice = OriginalPrice * (1 - MarkdownPercent / 100)
End Function
