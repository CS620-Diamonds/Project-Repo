Sub PopulateDaysOffIndex(daysOffIndex As Variant)
    ReDim daysOffIndex(5) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 4
        daysOffIndex(i) = Worksheets("IndexValues").Range("G" & (i + 2))
    Next i
End Sub
