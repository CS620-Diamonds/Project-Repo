Sub Initialize()
    On Error Resume Next
    Err.Clear
    Call DeclareGlobals
    Call PopulateFacultyIndex(facultyIndex)
    Call PopulateStatusIndex(statusIndex)
    Call PopulateGradUndergradIndex(gradUndergradIndex)
    Call PopulateCourseIndex(courseIndex)
    Call PopulateDaysOffFlag(daysOffFlag)
    Call PopulateDaysOffIndex(daysOffIndex)
    Call PopulateTimeIndex(timeBlockIndex)
    Call PopulateB2BFlag(back2backFlag)
    If Err.Number = 0 Then
        MsgBox "Worksheet is successfully initialized: Please press 'Create Schedule' to proceed"
    End If
    If Err.Number <> 0 Then
        MsgBox "There was an error populating the index arrays."
    End If
End Sub
