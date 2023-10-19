Sub CopyImagesFromPaths()
    Dim TargetRange As Range
    Dim Cell As Range
    Dim ImgPath As String
    Dim ImgTop As Double
    Dim ImgLeft As Double
    Dim ImgWidth As Double
    Dim ImgHeight As Double
    Dim Img As Object
    
    ' Define the target range where image paths are located
    Set TargetRange = Application.InputBox("Select the range containing image paths:", Type:=8)
    
    ' Exit if the user cancels the input box
    If TargetRange Is Nothing Then Exit Sub
    
    ' Loop through each cell in the selected range
    For Each Cell In TargetRange
        ImgPath = Cell.Value ' Get the image path from the cell
        
        ' Check if the cell contains a valid image path
        If ImgPath <> "" And Len(Dir(ImgPath)) > 0 Then
            ' Calculate the top, left, width, and height for the image
            ImgTop = Cell.Top
            ImgLeft = Cell.Left
            ImgWidth = Cell.Width
            ImgHeight = Cell.Height
            
            ' Insert the image into the active sheet
            Set Img = ActiveSheet.Pictures.Insert(ImgPath)
            With Img
                .Top = ImgTop
                .Left = ImgLeft
                .Width = ImgWidth
                .Height = ImgHeight
            End With
        End If
    Next Cell
End Sub
