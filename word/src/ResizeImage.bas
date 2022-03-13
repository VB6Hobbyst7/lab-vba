Sub ResizeImage(control As IRibbonControl)
'   Purpose: Resize selected image
'   Source: https://www.extendoffice.com/documents/word/1207-word-resize-all-multiple-images.html

    Dim shp As Word.Shape
    Dim ishp As Word.InlineShape
    If Word.Selection.Type <> wdSelectionInlineShape And _
        Word.Selection.Type <> wdSelectionShape Then
            Exit Sub
    End If
    If Word.Selection.Type = wdSelectionInlineShape Then
        Set ishp = Word.Selection.Range.InlineShapes(1)
        ishp.LockAspectRatio = False
        ishp.Height = CentimetersToPoints(5)
        ishp.Width = CentimetersToPoints(5)
    Else
        If Word.Selection.Type = wdSelectionShape Then
            Set shp = Word.Selection.ShapeRange(1)
            shp.LockAspectRatio = False
            shp.Height = CentimetersToPoints(5)
            shp.Width = CentimetersToPoints(5)
        End If
    End If
    
End Sub

