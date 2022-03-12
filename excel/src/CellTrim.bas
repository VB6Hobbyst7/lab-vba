Sub CellTrim(control As IRibbonControl)
'   Purpose: Trim spaces in cell
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
        For Each cell In rng
    cell.Value = Trim(cell)
    Next cell
    
ErrorHandler:
    Exit Sub

End Sub

