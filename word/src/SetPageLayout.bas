Sub SetPageLayout(control As IRibbonControl)
'   Purpose: Set page margin and edge distance

    With ActiveDocument.PageSetup
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(3)
        .RightMargin = CentimetersToPoints(1.8)
        .HeaderDistance = CentimetersToPoints(1)
        .FooterDistance = CentimetersToPoints(1)
        .PaperSize = wdPaperA4
    End With

End Sub

