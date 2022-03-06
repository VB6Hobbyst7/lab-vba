Sub FontArial()
'   Purpose: Set selected range to Arial
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim rng As Range
    Set rng = Selection
    rng.Font.Name = "Arial"
    rng.Font.Size = 10

ErrorHandler:
    Exit Sub

End Sub