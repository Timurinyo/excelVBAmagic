# excelVBAmagic

## Description
This repository holds some VBA scripts for Excel that saved me tons of time at least once. Hopefully they will help someone else or me in the future.

## Scripts
### ConvertCellsToShapes
```vba
'Converts cells in the selection into draggable shapes.
Sub ConvertCellsToShapes()
    Dim cel As Range
    Dim selectedRange As Range
    Dim Shp As Shape

    Set selectedRange = Application.Selection

    For Each cel In selectedRange.Cells
        If cel.Value <> "" Then
            Debug.Print cel.Address, cel.Value
            Set Shp = ActiveSheet.Shapes.AddShape(1, cel.Left, cel.Top, cel.Width, cel.Height)
            Shp.TextFrame.Characters.Text = cel.Value
            Shp.Fill.ForeColor.RGB = cel.Interior.Color
            Shp.TextFrame.Characters.Font.Color = cel.Font.Color
        End If
    Next cel

End Sub
```

### DeleteAllShapes
```vba
'Deletes all shapes from sheet
Sub DeleteAllShapes()
    Dim Shp As Shape
    
    For Each Shp In ActiveSheet.Shapes
        If Shp.Type = msoAutoShape Or Shp.Type = msoTextBox Then Shp.Delete
    Next Shp

End Sub
```

### ColorWhiteNonEmptyCells
```vba
'Colors white non-empty cells in selection into skyblue color
Sub ColorWhiteNonEmptyCells()
    Dim cel As Range
    Dim selectedRange As Range

    Set selectedRange = Application.Selection

    For Each cel In selectedRange.Cells
        If cel.Value <> "" And cel.Interior.Color = RGB(255, 255, 255) Then
            cel.Interior.Color = RGB(106, 173, 254)
        End If
    Next cel
End Sub
```
