Sub WorkbookFormulaToValue()
'   Purpose: Convert all workbook formulas to values (most efficient way)
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Dim sh As Worksheet, HidShts As New Collection

    For Each sh In ActiveWorkbook.Worksheets
        If Not sh.Visible Then
            HidShts.Add sh
            sh.Visible = xlSheetVisible
        End If
    Next sh
     
    Worksheets.Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
     
    For Each sh In HidShts
        sh.Visible = xlSheetHidden
    Next sh
   
ErrorHandler:
    Exit Sub

End Sub