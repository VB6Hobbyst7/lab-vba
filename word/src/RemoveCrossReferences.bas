Sub RemoveCrossReferences(control As IRibbonControl)
'   Purpose: Remove all cross-references

    Dim fld As Field
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldRef Then
            fld.Unlink
        End If
    Next
 
End Sub

