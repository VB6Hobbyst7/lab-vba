Sub ResetObject(control As IRibbonControl)
'   Purpose: Reset WorkbookObject sizes

    Dim shp As InlineShape
    
    For Each shp In ActiveDocument.InlineShapes
        With shp
            .LockAspectRatio = msoFalse
            .Reset
        End With
    Next shp
    
End Sub

