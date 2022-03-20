Sub WorkbookFont()
Sub SheetColumnsFS()
Sub SheetColumnsNTA4X()
Sub SheetColumnsWP()
Sub InsertWorkdone()
Sub InsertColumnWidth()
Sub InsertArrowDown()
Sub SelectBold()
Sub CellTrim()
Sub CaseUpper()
Sub CaseProper()
Sub CaseSentence()
Sub FormatTextToValue()
Sub FormatFontBlue()
Sub FormatFontGreen()
Sub FormatCellRed()
Sub FormatHighlightRed()
Sub FormatHighlightGreen()
Sub FormulaRound()
Sub FormulaThousands()
Sub FormatAccounting()
Sub FormulaToValue()
Sub FormulaReverseSign()
Sub RemoveBlankRows()
Sub SheetFontGeorgia()
Sub SheetFontArial()
Sub SheetFontSize8()
Sub SheetFontSize10()
Sub SheetFormulaToValue()
Sub SheetRemoveBlankRows()
Sub SheetPageBreakOff()
Sub InsertCrossReference()


' FormatHighlightGreen
' rng.Interior.Color = RGB(204, 285, 204)
' FormatHighlightRed()
' rng.Interior.Color = RGB(255, 204, 204)
' FormatCellRed()
' rng.Interior.Color = RGB(122, 24, 24)
' rng.Font.Color = RGB(255, 255, 255)
' FormatFontGreen()
' rng.Font.Color = RGB(0, 176, 80)
' FormatFontBlue()
' rng.Font.Color = RGB(0, 112, 192)
' Sub SelectBold()
' Dim rng As Range
' Dim WorkRng As Range
' Dim OutRng As Range
' On Error Resume Next
' xTitleId = "KutoolsforExcel"
' Set WorkRng = Application.Selection
' Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
' For Each rng In WorkRng
' If rng.Font.Bold Then
' If OutRng Is Nothing Then
' Set OutRng = rng
' Else
' Set OutRng = Union(OutRng, rng)
' End If
' End If
' Next
' If Not OutRng Is Nothing Then
' OutRng.Select
' End If
' End Sub



' rng.Value = "Workdone:"
' rng.Font.Bold = True
' rng.Offset(1, 0) = "TB"
' rng.Offset(1, 1) = ": Agreed to current year trial balance."
' rng.Offset(2, 0) = "PY"
' rng.Offset(2, 1) = ": Agreed to prior year audited balance."
' rng.Offset(3, 0) = "imm"
' rng.Offset(3, 1) = ": Immaterial (below CTT), suggest to leave."
' rng.Offset(4, 0) = "^"
' rng.Offset(4, 1) = ": Casted."
' rng.Offset(1, 0).Characters(1, 3).Font.Color = RGB(0, 112, 192)
' rng.Offset(2, 0).Characters(1, 3).Font.Color = RGB(255, 51, 0)
' rng.Offset(3, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
' rng.Offset(4, 0).Characters(1, 3).Font.Color = RGB(0, 176, 80)
' rng.Offset(1, 0).Characters(1, 3).Font.Bold = True
' rng.Offset(2, 0).Characters(1, 3).Font.Bold = True
' rng.Offset(3, 0).Characters(1, 3).Font.Bold = True
' rng.Offset(4, 0).Characters(1, 3).Font.Bold = True