Sub CaseProper()
'   Purpose: Set upper case on selection
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
    For Each Cell In rng
        Cell.Value = StrConv(Cell, vbProperCase)
    Next Cell
    
ErrorHandler:
    Exit Sub

End Sub
