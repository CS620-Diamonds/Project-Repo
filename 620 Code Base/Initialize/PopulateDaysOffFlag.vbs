Sub PopulateDaysOffFlag(daysOffFlag As Variant)
    ReDim daysOffFlag(2) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 1
        daysOffFlag(i) = Worksheets("IndexValues").Range("F" & (i + 2))
    Next i
End Sub
