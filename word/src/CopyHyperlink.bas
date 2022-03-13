Sub CopyHyperlink(control As IRibbonControl)
'   Purpose: Copy hyperlinks
'   Reference: https://www.msofficeforums.com/word-vba/38223-how-extract-selected-hyperlink-address-clipboard.html
'   Reference: https://software-solutions-online.com/word-vba-move-cursor-to-end-of-document/
'   Reference: https://gregmaxey.com/word_tips.html
'   Reference: https://www.thespreadsheetguru.com/blog/dynamically-populating-array-vba-variables
'   Reference: https://stackoverflow.com/questions/39690078/vba-output-contents-of-array-to-word-document
'   =========================================
'    For i = 1 To Selection.Hyperlinks.Count
'        With Selection.Hyperlinks(i)
'          StrTxt = .Address
'          If .SubAddress <> "" Then StrTxt = StrTxt & "#" & .SubAddress
'          With .Range.Fields(1).Code
'            .Text = StrTxt
'            .Copy
'          End With
'        End With
'        ActiveDocument.Undo
'    Next
'    Selection.EndKey Unit:=wdStory
'    Selection.Range.Text = vbNewLine
'    Selection.Paste

    Dim StrTxt As String
    Dim results() As Variant
    Dim inputWord As Variant
    Dim i As Long
    Dim insertPos As Range
    Set insertPos = Selection.Range
    
    ReDim results(Selection.Hyperlinks.Count)
    For i = 1 To Selection.Hyperlinks.Count
        results(i) = Selection.Hyperlinks(i).TextToDisplay & ": " & Selection.Hyperlinks(i).Address
    Next

    For Each inputWord In results
        insertPos.Collapse wdCollapseEnd
        insertPos = inputWord & vbCrLf
    Next
     
End Sub

