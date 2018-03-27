Sub PopulateGradUndergradIndex(gradUndergradIndex As Variant)
    ReDim gradUndergradIndex(3, 2) As Variant
    'index values
        gradUndergradIndex(0, 0) = 1
        gradUndergradIndex(1, 0) = 2
        gradUndergradIndex(2, 0) = 3
    'Grad/Undergrad Preference
        gradUndergradIndex(0, 1) = "Undergraduate Only"
        gradUndergradIndex(1, 1) = "Graduate Only"
        gradUndergradIndex(2, 1) = "No Preference"
End Sub
