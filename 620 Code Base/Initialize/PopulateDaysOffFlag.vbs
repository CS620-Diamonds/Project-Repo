Sub PopulateDaysOffFlag(daysOffFlag As Variant)
    ReDim daysOffFlag(2, 2) As Variant
    'index values
        daysOffFlag(0, 0) = 0
        daysOffFlag(1, 0) = 1
    'back 2 back flag
        daysOffFlag(0, 1) = "Yes"
        daysOffFlag(1, 1) = "No"
End Sub
