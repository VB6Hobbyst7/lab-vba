Function XSPELLNUMBER(ByVal MyNumber)
'   Purpose: Spell numbers as dollars
'   Source: https://support.microsoft.com/en-us/office/convert-numbers-into-words-a0d166fb-e1ea-4090-95c8-69442cd55d98
'   Reference: https://stackoverflow.com/questions/11155912/how-to-make-vba-function-vba-only-and-disable-it-as-udf/41130822
'   Modification:
'   - Made private GetHundreds, GetTens. Users do not need to access these.
'   - Added "Only" to the end of result.
'   - Fixed additional spacing between words in the original code.
'   - Fixed 0.XX appearing as "No Dollar and XX Cents" to "XX Cents Only"
'   Updated: 2022FEB23

'   Saves workbook before macro changes
    ActiveWorkbook.Save
    
    Dim Dollars, Cents, Temp
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "

'   String representation of amount
    MyNumber = Trim(Str(MyNumber))
    
'   Position of decimal place 0 if none
    DecimalPlace = InStr(MyNumber, ".")
    
'   Convert cents and set MyNumber to dollar amount
    If DecimalPlace > 0 Then
        Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    
    Count = 1
    Do While MyNumber <> ""
        Temp = GetHundreds(Right(MyNumber, 3))
        If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
            If Len(MyNumber) > 3 Then
                MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    
    Select Case Dollars
        Case "": Dollars = "" ' Originally "No Dollar"
        Case "One": Dollars = "One Dollar"
        Case Else
        Dollars = Dollars & " Dollars"
    End Select
    
    Select Case Cents
        Case "": Cents = " Only"
        Case "One": Cents = " and One Cent Only"
        Case Else
        Cents = " and " & Cents & " Cents Only"
    End Select
    
    If Dollars = "" Then
        XSPELLNUMBER = Trim(Replace(Replace(Dollars & Cents, "  ", " "), "and", ""))
    Else
        XSPELLNUMBER = Trim(Replace(Dollars & Cents, "  ", " "))
    End If

End Function
       
Private Function GetHundreds(ByVal MyNumber)
'   Purpose: Converts a number from 100-999 into text

    Dim Result As String
    
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    
'   Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
    End If
    
'   Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    
    GetHundreds = Result

End Function
    
Private Function GetTens(TensText)
'   Purpose: Converts a number from 10 to 99 into text.
 
    Dim Result As String
    
'   Null out the temporary function value.
    Result = ""
    
'   If value between 10-19...
    If Val(Left(TensText, 1)) = 1 Then
        Select Case Val(TensText)
            Case 10: Result = "Ten"
            Case 11: Result = "Eleven"
            Case 12: Result = "Twelve"
            Case 13: Result = "Thirteen"
            Case 14: Result = "Fourteen"
            Case 15: Result = "Fifteen"
            Case 16: Result = "Sixteen"
            Case 17: Result = "Seventeen"
            Case 18: Result = "Eighteen"
            Case 19: Result = "Nineteen"
            Case Else
        End Select
    Else
'   If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "Twenty"
            Case 3: Result = "Thirty"
            Case 4: Result = "Forty"
            Case 5: Result = "Fifty"
            Case 6: Result = "Sixty"
            Case 7: Result = "Seventy"
            Case 8: Result = "Eighty"
            Case 9: Result = "Ninety"
            Case Else
        End Select
'   Retrieve ones place.
        Result = Result & " " & GetDigit(Right(TensText, 1))
    End If
  
    GetTens = Result
 
End Function
  
Private Function GetDigit(Digit)
'   Purpose: Converts a number from 1 to 9 into text.

    Select Case Val(Digit)
        Case 1: GetDigit = "One"
        Case 2: GetDigit = "Two"
        Case 3: GetDigit = "Three"
        Case 4: GetDigit = "Four"
        Case 5: GetDigit = "Five"
        Case 6: GetDigit = "Six"
        Case 7: GetDigit = "Seven"
        Case 8: GetDigit = "Eight"
        Case 9: GetDigit = "Nine"
        Case Else: GetDigit = ""
    End Select
    
End Function
