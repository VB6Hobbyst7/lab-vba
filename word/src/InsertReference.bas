Sub InsertReference(control As IRibbonControl)
'   Purpose: Paste clipboard content as hyperlink
'   References: https://www.slipstick.com/developer/code-samples/paste-clipboard-contents-vba/
'   Notes:
'   - https://excel-macro.tutorialhorizon.com/vba-excel-reference-libraries-in-excel-workbook/

    Dim DataObj As MSForms.DataObject
    Set DataObj = New MSForms.DataObject
    Dim strPaste As Variant
    DataObj.GetFromClipboard
    
    Application.ScreenUpdating = False
    
    strPaste = DataObj.GetText(1)
    If strPaste = False Then Exit Sub
    If strPaste = "" Then Exit Sub

    Selection.TypeText Text:="["
    ActiveDocument.Hyperlinks.Add _
        Anchor:=Selection.Range, _
        Address:=strPaste, _
        TextToDisplay:=ChrW(664)
    Selection.TypeText Text:="]"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Color = wdColorBlue
    Selection.Move
    Selection.MoveRight Unit:=wdCharacter, Count:=1

    Set DataObj = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

