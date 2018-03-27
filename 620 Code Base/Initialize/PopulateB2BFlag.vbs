Sub PopulateB2BFlag(back2backFlag As Variant)
    ReDim back2backFlag(2) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 1
        back2backFlag(i) = Worksheets("IndexValues").Range("I" & (i + 2))
    Next i
End Sub
