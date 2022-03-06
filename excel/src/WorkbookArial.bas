Sub WorkbookArial()
'   Purpose: Standardise workbook font type and size
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    For Each ws In Worksheets
         With ws
            .Cells.Font.Name = "Arial"
            .Cells.Font.Size = 10
         End With
    Next ws
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Activate
        ActiveWindow.Zoom = 100
    Next
    Application.ScreenUpdating = True

ErrorHandler:
    Exit Sub

End Sub