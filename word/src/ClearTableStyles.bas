Sub ClearTableStyles(control As IRibbonControl)
'   Purpose: Clear table styles

    Dim objTable As Table
    Dim objDoc As Document
    
    Application.ScreenUpdating = False
    Set objDoc = ActiveDocument
    
    For Each objTable In objDoc.Tables
      objTable.Style = "Table Normal"
      objTable.Borders.Enable = True
    Next objTable
    
    Application.ScreenUpdating = True
    
    Set objDoc = Nothing
  
End Sub

