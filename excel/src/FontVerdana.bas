Sub FontVerdana()
'   Purpose: Set selected range to Verdana
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
    rng.Font.Name = "Verdana"
    rng.Font.Size = 10

ErrorHandler:
    Exit Sub

End Sub