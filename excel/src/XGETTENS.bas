Private Function XGETTENS(TensText)
'   Purpose: Converts a number from 10 to 99 into text.
 
    Dim result As String
    
'   Null out the temporary function value.
    result = ""
    
'   If value between 10-19...
    If Val(Left(TensText, 1)) = 1 Then
        Select Case Val(TensText)
            Case 10: result = "Ten"
            Case 11: result = "Eleven"
            Case 12: result = "Twelve"
            Case 13: result = "Thirteen"
            Case 14: result = "Fourteen"
            Case 15: result = "Fifteen"
            Case 16: result = "Sixteen"
            Case 17: result = "Seventeen"
            Case 18: result = "Eighteen"
            Case 19: result = "Nineteen"
            Case Else
        End Select
    Else
'   If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: result = "Twenty"
            Case 3: result = "Thirty"
            Case 4: result = "Forty"
            Case 5: result = "Fifty"
            Case 6: result = "Sixty"
            Case 7: result = "Seventy"
            Case 8: result = "Eighty"
            Case 9: result = "Ninety"
            Case Else
        End Select
'   Retrieve ones place.
        result = result & " " & XGETDIGIT(Right(TensText, 1))
    End If
  
    XGETTENS = result
 
End Function

