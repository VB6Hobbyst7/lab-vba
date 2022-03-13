Sub SplitVertically(control As IRibbonControl)
'   Purpose: Split WORD active window vertically to view side by side
'   Source: https://dharma-records.buddhasasana.net/computing/ms-word-split-windows-vertically
 
    Dim Win1 As Integer
    Dim Win2 As Integer
     
    Dim WinWidth As Integer
    Dim WinHeight As Integer
 
'   Check for duplicated window
 
    Dim Win As Word.Window
    Dim DocString As String
    Dim FirstString
    Dim SecondString
    Dim StringLength As Long
     
    For Each Win In Application.Windows
    DocString = Win
    FirstString = Right(DocString, 1)
     
        If FirstString = "1" Then
         
'   If there is a duplicate
         
        StringLength = Len(DocString) - 2
        SecondString = Left(DocString, StringLength)
         
'   Close the copy
     
        Windows(SecondString & ":2").Close
         
'   Activate and maximise the identified document window
     
        Windows(Win).Activate
        Windows(Win).WindowState = wdWindowStateMaximize
         
        GoTo TheEnd
         
        End If

'   Otherwise check the next window
     
    Next Win
     
'   If there are no duplicates, get the dimensions
     
    ActiveWindow.WindowState = wdWindowStateMaximize
     
'   Find the serial number of the window and set the variable
     
    Win1 = ActiveWindow.Index
     
'   Set the dimension variables
     
    WinHeight = ActiveWindow.Height - 20
    WinWidth = ActiveWindow.Width

'   Make a new window from the first
     
    NewWindow
     
'   Find the serial number of the new window
         
    Win2 = ActiveWindow.Index
         
'   Arrange all windows (window must be in maximised state)
     
    Windows.Arrange
     
'   Set the size of the two windows we found
     
    With Windows(Win1)
        .Left = 0
        .Top = 0
        .Height = WinHeight
        .Width = WinWidth / 2
    End With
     
    With Windows(Win2)
        .Left = WinWidth / 2
        .Top = 0
        .Height = WinHeight
        .Width = WinWidth / 2
    End With
        
'   Return to the first window
     
    Windows(Win1).Activate
         
'   Added by Ryan
'   Resize windows
'   Width = 486

    Application.Resize Width:=WinWidth / 2, Height:=Application.UsableHeight
    Windows(Win2).Activate
    Application.Resize Width:=WinWidth / 2, Height:=Application.UsableHeight
    
TheEnd:
         
End Sub

