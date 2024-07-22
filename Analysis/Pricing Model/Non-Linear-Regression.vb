Option Explicit

' Function to calculate profit margin after accounting for shrink cost
Function AdjustedMargin(Sales As Double, Cost As Double, ShrinkCost As Double) As Double
    If Sales = 0 Then
        AdjustedMargin = 0
    Else
        AdjustedMargin = (Sales - (Cost + ShrinkCost)) / Sales
    End If
End Function

' Function to perform non-linear regression (simplified for demonstration)
Sub NonLinearRegression()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Dim LastRow As Long
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    Dim Sales() As Double
    Dim Cost() As Double
    Dim ShrinkCost() As Double
    Dim Units() As Double
    
    ReDim Sales(1 To LastRow - 1)
    ReDim Cost(1 To LastRow - 1)
    ReDim ShrinkCost(1 To LastRow - 1)
    ReDim Units(1 To LastRow - 1)
    
    ' Load data into arrays
    For i = 2 To LastRow
        Sales(i - 1) = ws.Cells(i, 4).Value
        Cost(i - 1) = ws.Cells(i, 5).Value
        ShrinkCost(i - 1) = ws.Cells(i, 7).Value
        Units(i - 1) = ws.Cells(i, 3).Value
    Next i
    
    ' Perform non-linear regression (simplified example)
    ' Here, we'll just demonstrate using a simple polynomial fit
    Dim Coefficients() As Double
    Coefficients = PolynomialFit(Sales, Cost, ShrinkCost, Units, 2)
    
    ' Output results
    ws.Cells(1, 9).Value = "Coefficient A"
    ws.Cells(1, 10).Value = "Coefficient B"
    ws.Cells(1, 11).Value = "Coefficient C"
    ws.Cells(2, 9).Value = Coefficients(0)
    ws.Cells(2, 10).Value = Coefficients(1)
    ws.Cells(2, 11).Value = Coefficients(2)
    
    ' Calculate adjusted margins
    ws.Cells(1, 12).Value = "Adjusted Margin"
    For i = 2 To LastRow
        ws.Cells(i, 12).Value = AdjustedMargin(Sales(i - 1), Cost(i - 1), ShrinkCost(i - 1))
    Next i
End Sub

' Polynomial fit function (simplified for demonstration)
Function PolynomialFit(Sales() As Double, Cost() As Double, ShrinkCost() As Double, Units() As Double, Degree As Integer) As Double()
    Dim Coefficients(0 To 2) As Double
    ' Simplified example of fitting a polynomial to data
    ' In practice, use a robust method such as Gauss-Newton or Levenberg-Marquardt algorithm
    
    ' Dummy coefficients for demonstration
    Coefficients(0) = 0.5
    Coefficients(1) = 0.3
    Coefficients(2) = 0.2
    
    PolynomialFit = Coefficients
End Function
