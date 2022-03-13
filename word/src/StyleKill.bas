Sub StyleKill(control As IRibbonControl)
'   Purpose: Delete unwanted styles
'   Source: https://word.tips.net/T001337_Removing_Unused_Styles.html

    Dim oStyle As Style
    For Each oStyle In ActiveDocument.Styles
        'Only check out non-built-in styles
        If oStyle.BuiltIn = False Then
                oStyle.Delete
        End If
    Next oStyle
     
End Sub

