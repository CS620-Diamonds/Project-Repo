Sub PopulateStatusIndex(statusIndex As Variant)
    ReDim statusIndex(3, 2) As Variant
    'index values
        statusIndex(0, 0) = 1
        statusIndex(1, 0) = 2
    'status values
        statusIndex(0, 1) = "Part Time"
        statusIndex(1, 1) = "Full Time"
    'max classes for given status
        statusIndex(0, 2) = 2
        statusIndex(1, 2) = 4
End Sub
