Sub SheetColumnsFS()
'   Purpose: Standardise columns width for specific worksheet: BS/PL tab
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Columns.COLUMNWIDTH = 14
    Columns("A").COLUMNWIDTH = 3
    Columns("B").COLUMNWIDTH = 1
    Columns("C").COLUMNWIDTH = 28
    Columns("D").COLUMNWIDTH = 4
    Columns("E").COLUMNWIDTH = 11
    Columns("F").COLUMNWIDTH = 1
    Columns("G").COLUMNWIDTH = 11
    Columns("H").COLUMNWIDTH = 1
    Columns("I").COLUMNWIDTH = 11
    Columns("J").COLUMNWIDTH = 1
    Columns("K").COLUMNWIDTH = 11
    Columns("L").COLUMNWIDTH = 1
    Columns("M").COLUMNWIDTH = 1
    Columns("N").COLUMNWIDTH = 1
    Columns("M").Interior.Color = RGB(217, 217, 217)
    Columns("A:L").Font.Name = "Times New Roman"
    Columns("A:L").Font.Size = 10
    Range("B1").Formula = "=XCOLUMNWIDTH(B1)"
    Range("C1").Formula = "=XCOLUMNWIDTH(C1)"
    Range("D1").Formula = "=XCOLUMNWIDTH(D1)"
    Range("E1").Formula = "=XCOLUMNWIDTH(E1)"
    Range("F1").Formula = "=XCOLUMNWIDTH(F1)"
    Range("G1").Formula = "=XCOLUMNWIDTH(G1)"
    Range("H1").Formula = "=XCOLUMNWIDTH(H1)"
    Range("I1").Formula = "=XCOLUMNWIDTH(I1)"
    Range("J1").Formula = "=XCOLUMNWIDTH(J1)"
    Range("K1").Formula = "=XCOLUMNWIDTH(K1)"
    Range("O1").Formula = "=SUM(B1:K1)"
    Range("B1").HorizontalAlignment = xlCenter
    Range("C1").HorizontalAlignment = xlCenter
    Range("D1").HorizontalAlignment = xlCenter
    Range("E1").HorizontalAlignment = xlCenter
    Range("F1").HorizontalAlignment = xlCenter
    Range("G1").HorizontalAlignment = xlCenter
    Range("H1").HorizontalAlignment = xlCenter
    Range("I1").HorizontalAlignment = xlCenter
    Range("J1").HorizontalAlignment = xlCenter
    Range("K1").HorizontalAlignment = xlCenter
    Range("O1").HorizontalAlignment = xlLeft
    
ErrorHandler:
    Exit Sub

End Sub