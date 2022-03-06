Sub SheetColumnsNTA4X()
'   Purpose: Standardise columns width for specific worksheet: NTA 4-columns tab
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Columns.COLUMNWIDTH = 14
    Columns("A").COLUMNWIDTH = 3
    Columns("B").COLUMNWIDTH = 1
    Columns("C").COLUMNWIDTH = 14
    Columns("D").COLUMNWIDTH = 1
    Columns("E").COLUMNWIDTH = 13
    Columns("F").COLUMNWIDTH = 1
    Columns("G").COLUMNWIDTH = 10
    Columns("H").COLUMNWIDTH = 1
    Columns("I").COLUMNWIDTH = 10
    Columns("J").COLUMNWIDTH = 1
    Columns("K").COLUMNWIDTH = 10
    Columns("L").COLUMNWIDTH = 1
    Columns("M").COLUMNWIDTH = 10
    Columns("N").COLUMNWIDTH = 1
    Columns("O").COLUMNWIDTH = 1
    Columns("P").COLUMNWIDTH = 1
    Columns("O").Interior.Color = RGB(217, 217, 217)
    Columns("A:N").Font.Name = "Times New Roman"
    Columns("A:N").Font.Size = 10
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
    Range("L1").Formula = "=XCOLUMNWIDTH(L1)"
    Range("M1").Formula = "=XCOLUMNWIDTH(M1)"
    Range("Q1").Formula = "=SUM(B1:M1)"
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
    Range("L1").HorizontalAlignment = xlCenter
    Range("M1").HorizontalAlignment = xlCenter
    Range("Q1").HorizontalAlignment = xlLeft

ErrorHandler:
    Exit Sub

End Sub