Sub SetTablesBordersColor(control As IRibbonControl, varColor As Long)
'   Purpose: Standardise all table borders color in a document at 1/2 pt single-line

    Application.ScreenUpdating = False
    
    Dim Tbl As Table
    For Each Tbl In ActiveDocument.Tables
        Tbl.Borders.InsideColor = varColor
        Tbl.Borders.InsideLineStyle = wdLineStyleSingle
'        Tbl.Borders.InsideLineWidth = wdLineWidth050pt
        Tbl.Borders.OutsideColor = varColor
        Tbl.Borders.OutsideLineStyle = wdLineStyleSingle
'        Tbl.Borders.OutsideLineWidth = wdLineWidth050pt
    Next
        
    Application.ScreenUpdating = True
    Application.ScreenRefresh

End Sub

