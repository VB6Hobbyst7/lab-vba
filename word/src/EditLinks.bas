Sub EditLinks(control As IRibbonControl)
'   Purpose: Edit hyperlinks
'   Reference: https://stackoverflow.com/questions/3355266/how-to-programmatically-edit-all-hyperlinks-in-a-word-document
'   Reference: http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.hyperlink_members.aspx
'   Reference: https://shaunakelly.com/word/word-development/selecting-or-referring-to-a-page-in-the-word-object-model.html
'   =========================================
'   doc.Hyperlinks(i).Address = Replace(doc.Hyperlinks(i).Address, "gopher:", "https://")
'   If LCase(doc.Hyperlinks(i).Address) Like "*partOfHyperlinkHere*" Then
'   doc.Hyperlinks(i).Address = Mid(doc.Hyperlinks(i).Address, 70,20)

    Dim i As Long
    For i = 1 To Selection.Hyperlinks.Count
        Selection.Hyperlinks(i).TextToDisplay = "[" & Selection.Hyperlinks(i).TextToDisplay & "]"
    Next
    
    Call CopyHyperlink(control)
    
End Sub

