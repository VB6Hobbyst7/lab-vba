Sub KillTheHyperlinks(control As IRibbonControl)
'   Purpose: Removes all hyperlinks from the document

    With ThisDocument
        While .Hyperlinks.Count > 0
            .Hyperlinks(1).Delete
        Wend
    End With
    Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = False
    
End Sub

