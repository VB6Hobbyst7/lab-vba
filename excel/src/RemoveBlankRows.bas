Sub RemoveBlankRows()
'   Purpose: Remove blank rows in selection
'   Reference: https://www.wallstreetmojo.com/vba-last-row/
'   Updated: 2022FEB25

'   Saves workbook before macro changes
    On Error GoTo ErrorHandler
    ActiveWorkbook.Save

    Application.ScreenUpdating = False
        
    Dim SourceRange As Range
    Dim lastRow As Long
    
    Set SourceRange = Selection

    lastRow = SourceRange.SpecialCells(xlCellTypeLastCell).row

    If Not (SourceRange Is Nothing) Then
        For i = lastRow To 1 Step -1
            Set EntireRow = SourceRange.Cells(i, 1).EntireRow
            If Application.WorksheetFunction.CountA(EntireRow) = 0 Then
                EntireRow.Delete
            End If
        Next
    End If

    Application.ScreenUpdating = True
    
ErrorHandler:
    Exit Sub

End Sub