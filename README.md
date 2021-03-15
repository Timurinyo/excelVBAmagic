# excelVBAmagic

## Description
This repository holds some VBA scripts for Excel that saved me tons of time. Hopefully they will help someone else or me in the future.

## Scripts
### ConvertCellsToShapes
'Converts cells in the selection into draggable shapes.
```vba
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
