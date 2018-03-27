Sub PopulateFacultyIndex(facultyIndex As Variant) '
    ReDim facultyIndex(31) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 31
        facultyIndex(i) = Worksheets("IndexValues").Range("B" & (i + 2))
    Next i
End Sub
