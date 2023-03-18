Sub ChangeLineWidth()

    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim border As Variant
    Dim NewWidth As Double
    
    ' Set the worksheet where you want to perform the line width replacement using the sheet index
    Set ws = ActiveSheet
    
    ' Set the range of cells where you want to perform the line width replacement
    Set rng = ws.Range("A1:BN72")
    
    ' Set the new line width
    NewWidth = 2 ' new line width (2 points)
    
    ' Loop through all the cells in the range
    For Each cell In rng
        For Each border In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
            With cell.Borders(border)
                ' If the border weight is greater than 1 point, change it to 2 points
                If .LineStyle <> xlLineStyleNone And .Weight > xlHairline + 1 Then
                    .Weight = xlMedium
                End If
            End With
        Next border
    Next cell

End Sub
