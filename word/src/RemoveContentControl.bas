Sub RemoveContentControl(control As IRibbonControl)
'   Purpose: Remove all content controls
    
    Dim oRng As Range
    Dim CC   As ContentControl
    Dim LC   As Integer
    Dim LRCC As Integer
    Dim LTCC As Integer
    Dim LE   As Boolean

    Set oRng = ActiveDocument.Content
    LTCC = LTCC + oRng.ContentControls.Count
    For LC = oRng.ContentControls.Count To 1 Step -1
    
    Set CC = oRng.ContentControls(LC)
    If CC.LockContentControl = True Then
        CC.LockContentControl = False
    End If
    CC.Delete
    If Not LE Then
        LRCC = LRCC + 1
        End If
        LE = False
    Next
    
End Sub

