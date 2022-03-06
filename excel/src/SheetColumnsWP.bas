Sub SheetColumnsWP()
'   Purpose: Standardise workbook columns width
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    For Each ws In Worksheets
        Columns.COLUMNWIDTH = 14
        Columns("A").COLUMNWIDTH = 1
        Columns("B").COLUMNWIDTH = 3
        Columns("C").COLUMNWIDTH = 5
    Next ws
    
ErrorHandler:
    Exit Sub

End Sub