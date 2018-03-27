Sub PopulateStatusIndex(statusIndex As Variant)
    ReDim statusIndex(2) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 1
        statusIndex(i) = Worksheets("IndexValues").Range("C" & (i + 2))
    Next i
End Sub
