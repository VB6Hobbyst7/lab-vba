Sub OpenEmbeddedExcelInWord(control As IRibbonControl)
'   Purpose: Remove LockAspectRatio from linked Excel objects

    Dim shp As InlineShape
    
    For Each shp In ActiveDocument.InlineShapes
        With shp
            .LockAspectRatio = msoFalse
            .Reset
        End With
    Next shp
    
End Sub

