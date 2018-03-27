Sub PopulateGradUndergradIndex(gradUndergradIndex As Variant)
    ReDim gradUndergradIndex(3) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 2
        gradUndergradIndex(i) = Worksheets("IndexValues").Range("D" & (i + 2))
    Next i
End Sub
