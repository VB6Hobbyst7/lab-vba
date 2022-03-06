Public Function XGCDISTANCE(textQuery As Range, varTarget As Range, varDictionary As Range)
'	Purpose: Fill array by criteria
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim result As Variant
    Dim trackResult As Variant
    Dim cLat As Double
    Dim cLon As Double
    Dim i As Single
    
    result = 0
    
    For i = 1 To varDictionary.Cells.Count
        If varDictionary.Cells(i, 1).Value = textQuery Then
            cLat = varDictionary.Cells(i, 2)
            cLon = varDictionary.Cells(i, 3)
            trackResult = Haversine(varTarget.Cells(1, 1), varTarget.Cells(1, 2), cLat, cLon)
            If result = 0 Then result = trackResult
            If trackResult < result Then result = trackResult
        End If
    Next
    
    XGCDISTANCE = result

End Function

Private Function Haversine(Lat1 As Variant, Lon1 As Variant, Lat2 As Variant, Lon2 As Variant)
 '	Purpose: Great Circle Distance calculation
 '	Note: Returns results in kilometers

    Dim R As Integer, dlon As Variant, dlat As Variant, Rad1 As Variant
    Dim a As Variant, c As Variant, d As Variant, Rad2 As Variant

    R = 6371
    dlon = Excel.WorksheetFunction.Radians(Lon2 - Lon1)
    dlat = Excel.WorksheetFunction.Radians(Lat2 - Lat1)
    Rad1 = Excel.WorksheetFunction.Radians(Lat1)
    Rad2 = Excel.WorksheetFunction.Radians(Lat2)
    a = Sin(dlat / 2) * Sin(dlat / 2) + Cos(Rad1) * Cos(Rad2) * Sin(dlon / 2) * Sin(dlon / 2)
    c = 2 * Excel.WorksheetFunction.Atan2(Sqr(1 - a), Sqr(a))
    d = R * c
    Haversine = d
	
End Function