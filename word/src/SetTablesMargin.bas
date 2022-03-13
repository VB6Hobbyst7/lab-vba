Sub SetTablesMargin(control As IRibbonControl, varPadding As Double)
'   Purpose: Standardise all table paddings in a document
'   varPadding: Measured in centimeters
'   Notes:

    Application.ScreenUpdating = False
    
    Dim Tbl As Table
    For Each Tbl In ActiveDocument.Tables
        Tbl.AutoFitBehavior (wdAutoFitWindow)
        Tbl.AllowAutoFit = True
        Tbl.LeftPadding = CentimetersToPoints(varPadding)
        Tbl.RightPadding = CentimetersToPoints(varPadding)
        Tbl.TopPadding = CentimetersToPoints(varPadding)
        Tbl.BottomPadding = CentimetersToPoints(varPadding)
    Next
    
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    
End Sub

