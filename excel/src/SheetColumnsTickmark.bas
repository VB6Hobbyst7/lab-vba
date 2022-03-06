Sub SheetColumnsTickmark()
'   Purpose: Standardise columns width for specific worksheet: Tickmark tab
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Columns.COLUMNWIDTH = 15
    Columns("A").COLUMNWIDTH = 3
    Columns("B").COLUMNWIDTH = 1
    Columns("C").COLUMNWIDTH = 3
    Columns("D").COLUMNWIDTH = 15
    Columns("E").COLUMNWIDTH = 15
    Columns("F").COLUMNWIDTH = 15
    Columns("G").COLUMNWIDTH = 15
    Columns("H").COLUMNWIDTH = 15
    Columns("I").COLUMNWIDTH = 15
    Columns("J").COLUMNWIDTH = 15
    Columns("K").COLUMNWIDTH = 15
    Columns("L").COLUMNWIDTH = 15
    Columns("M").COLUMNWIDTH = 1
    Columns("N").COLUMNWIDTH = 5
    
ErrorHandler:
    Exit Sub

End Sub