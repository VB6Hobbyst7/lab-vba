Sub Document_ContentControlOnExit(control As IRibbonControl, ByVal ContentControl As ContentControl, Cancel As Boolean)
'   Purpose: Change Textbox content as Dropdown List change.

    Dim oCC As ContentControl
    Dim oRng As Word.Range

    Select Case ContentControl.Title
      Case "Client" 'The "Client" Dropdown CC in document.
        Set oCC = ActiveDocument.SelectContentControlsByTitle("RegNo").Item(1) 'Richtext CC in Header
        Select Case True
          Case ContentControl.ShowingPlaceholderText:
            oCC.Range.Text = vbNullString
          Case ContentControl.Range.Text = "AC ALLIANCES (PAC)"
            oCC.Range.Text = "201118268H"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdRed
          Case ContentControl.Range.Text = "Y M WOO & CO"
            oCC.Range.Text = "S88PF0309G"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdBlue
          Case ContentControl.Range.Text = "GAAP PAC"
            oCC.Range.Text = "201831129C"
            Set oRng = oCC.Range
    '          oRng.Start = oRng.Start + 9
    '          oRng.Font.ColorIndex = wdGreen
        End Select
    End Select
lbl_Exit:
    Exit Sub
    
End Sub

