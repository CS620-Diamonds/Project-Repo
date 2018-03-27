Sub PopulateTimeIndex(timeBlockIndex As Variant)
    ReDim timeBlockIndex(26)
    Worksheets("IndexValues").Activate
    For i = 0 To 26
        timeBlockIndex(i) = Worksheets("IndexValues").Range("H" & (i + 2))
    Next i
End Sub
