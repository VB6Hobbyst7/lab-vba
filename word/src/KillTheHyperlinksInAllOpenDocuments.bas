Sub KillTheHyperlinksInAllOpenDocuments(control As IRibbonControl)
'   Purpose: Removes all hyperlinks from all opened document

    Dim doc As Document
    Dim szOpenDocName As String
     
    For Each doc In Application.Documents
        szOpenDocName = doc.Name
        With Documents(szOpenDocName)
            While .Hyperlinks.Count > 0
                .Hyperlinks(1).Delete
            Wend
        End With
        Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = False
    Next doc
    
End Sub

