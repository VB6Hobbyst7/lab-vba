Private Function HYPERACTIVE(ByRef rng As Range)
'   Purpose: To create hyperlink based on selected cell
'   Note: Passive function to be activated by InsertCrossReference()

    Dim strAddress As String
    Dim target As Range

    On Error Resume Next
        Set target = Application.InputBox( _
          Title:="Create Hyperlink", _
          Prompt:="Select a cell to create hyperlink", _
          Type:=8)
    On Error GoTo 0
  
'   Ensure User did not cancel
    If target Is Nothing Then Exit Function
  
'   Set Variable to first cell in user's input (ensuring only 1 cell)
    Set target = target.Cells(1, 1)
    
'   Get the text value of the address to display as hyperlink TextToDisplay
    strAddress = target.Parent.Name & "!" & target.Address(External:=False)

'   Generate hyperlink
    With ActiveSheet.Hyperlinks
    .Add Anchor:=rng, _
         Address:="", _
         SubAddress:="=" & strAddress, _
         TextToDisplay:=strAddress
    End With
    
End Function

