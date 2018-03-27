Sub PopulateDaysOffIndex(daysOffIndex As Variant)
    ReDim daysOffIndex(5, 2) As Variant
    'index values
        daysOffIndex(0, 0) = 1
        daysOffIndex(1, 0) = 2
        daysOffIndex(2, 0) = 3
        daysOffIndex(0, 0) = 4
        daysOffIndex(1, 0) = 5
    'days values
        daysOffIndex(0, 1) = "M"
        daysOffIndex(1, 1) = "T"
        daysOffIndex(2, 1) = "W"
        daysOffIndex(0, 1) = "R"
        daysOffIndex(1, 1) = "F"
End Sub
