Function XCOLUMNWIDTH(MR As Range) As Double
'   Purpose: Get column width
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Application.Volatile
    XCOLUMNWIDTH = MR.COLUMNWIDTH
    
End Function