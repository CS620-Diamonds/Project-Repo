Sub PopulateB2BFlag(back2backFlag As Variant)
    ReDim back2backFlag(2, 2) As Variant
        'index values
            back2backFlag(0, 0) = 0
            back2backFlag(1, 0) = 1
        'back 2 back flag
            back2backFlag(0, 1) = "Yes"
            back2backFlag(1, 1) = "No"
End Sub
