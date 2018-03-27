Sub PopulateCourseIndex(courseIndex As Variant)
    ReDim courseIndex(53) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 53
        courseIndex(i) = Worksheets("IndexValues").Range("E" & (i + 2))
    Next i
End Sub
