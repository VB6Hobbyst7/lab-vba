Sub InsertSymbol(control As IRibbonControl)
'   Purpose: Insert symbol for reference

    Application.ScreenUpdating = False
    Selection.InsertSymbol Font:="Arial", CharacterNumber:=664, Unicode:=True
    Application.ScreenUpdating = True
        
End Sub

