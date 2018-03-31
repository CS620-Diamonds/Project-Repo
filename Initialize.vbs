Sub Initialize()
    On Error Resume Next
    Err.Clear
    Call PopulateFacultyIndex(facultyIndex)
    Call PopulateStatusIndex(statusIndex)
    Call PopulateGradUndergradIndex(gradUndergradIndex)
    Call PopulateCourseIndex(courseIndex)
    Call PopulateDaysOffFlag(daysOffFlag)
    Call PopulateDaysOffIndex(daysOffIndex)
    Call PopulateTimeIndex(timeBlockIndex)
    Call PopulateB2BFlag(back2backFlag)
    Worksheets("Responses").Activate
    If Err.Number = 0 Then
        MsgBox "Worksheet is successfully initialized: Please press 'Create Schedule' to proceed"
    End If
    If Err.Number <> 0 Then
        MsgBox "There was an error populating the index arrays."
    End If
End Sub

Sub DeclareGlobals() 'Global Variable Instantiation
        Dim position
        Dim maxPerBlock As Integer
        maxPerBlock = 3

        Dim isFT As Boolean
        Dim priority As Integer 'faculty priority
        Dim status As String 'faculty status
        Dim maxClasses As Integer 'faculty max classes
        Dim bestAvailableTime(4) As Variant
        Dim bestAvailableCourse(4) As Variant
        Dim course As String
        Dim hasPrerequisite As Boolean
        Dim preRequisiteCourse As String
        Dim preRequisiteCourses(5) As String


    'Set delimiters
        Dim sectionDelim As String
        Dim courseDelim As String
        Dim timeDelim As String
        sectionDelim = "||"
        courseDelim = "{}"
        timeDelim = "[]"

    'Sheet Lenghts
        Dim TotalResponseRows As Integer 'instantiate and assign number of response rows
        Dim TotalDataImportRows As Integer
        Dim TotalScheduleRows As Integer

    'Object Arrays:
    'schedule
        Dim scheduleArray(1, 1) As Variant 'Instantiate schedule array - each index is a Faculty Object
        Dim scheduleFaculty(1, 1) As Variant 'Instantiate scheduleFaculty Object, from each index in ^
    'response
        Dim responseArray(1, 1) As Variant 'Instantiate response array - each index is a Faculty Object
        Dim responseFaculty(1, 1) As Variant 'Instantiate responseFaculty Object, from each index in ^
    'Priority
    Dim priorityArray(32) As String 'instantiate priority array, full of only that current priority level of resposnes

    'Index Arrays
        Dim timeBlockIndex(26) As Variant
        Dim timeNumberIndex(26) As Variant
        Dim courseIndex(53) As Variant
        Dim professorIndex(32) As Variant
        Dim statusIndex(2) As Variant
        Dim gradUndergradIndex(2) As Variant
        Dim daysOffFlag(2) As Variant
        Dim daysOffIndex(5) As Variant
        Dim back2backFlag(2) As Variant

    'Split Values
        Dim times(26) As Variant
        Dim courses(10) As String
        Dim i As Integer 'loop index
        Dim counter As Integer 'coursePreference Counter for assignment
        Dim strACAddress As Variant
        Dim strFaculty As String
        Dim timeCell

    'Get Sheet Lengths
        TotalResponseRows = Worksheets("Responses").Rows(Rows.Count).End(xlUp).Row
        'TotalDataImportRows = Worksheets("DataImport").Rows(Rows.Count).End(xlUp).Row
        'TotalScheduleRows = Worksheets("Schedule").Rows(Rows.Count).End(xlUp).Row

End Sub

Sub PopulateB2BFlag(back2backFlag As Variant)
    ReDim back2backFlag(2) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 1
        back2backFlag(i) = Worksheets("IndexValues").Range("I" & (i + 2))
    Next i
End Sub

Sub PopulateDaysOffFlag(daysOffFlag As Variant)
    ReDim daysOffFlag(2) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 1
        daysOffFlag(i) = Worksheets("IndexValues").Range("F" & (i + 2))
    Next i
End Sub

Sub PopulateDaysOffIndex(daysOffIndex As Variant)
    ReDim daysOffIndex(5) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 4
        daysOffIndex(i) = Worksheets("IndexValues").Range("G" & (i + 2))
    Next i
End Sub

Sub PopulateGradUndergradIndex(gradUndergradIndex As Variant)
    ReDim gradUndergradIndex(3) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 2
        gradUndergradIndex(i) = Worksheets("IndexValues").Range("D" & (i + 2))
    Next i
End Sub

Sub PopulateStatusIndex(statusIndex As Variant)
    ReDim statusIndex(2) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 1
        statusIndex(i) = Worksheets("IndexValues").Range("C" & (i + 2))
    Next i
End Sub

Sub PopulateFacultyIndex(facultyIndex As Variant) '
    ReDim facultyIndex(31) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 31
        facultyIndex(i) = Worksheets("IndexValues").Range("B" & (i + 2))
    Next i
End Sub

Sub PopulateCourseIndex(courseIndex As Variant)
    ReDim courseIndex(53) As Variant
    Worksheets("IndexValues").Activate
    For i = 0 To 53
        courseIndex(i) = Worksheets("IndexValues").Range("E" & (i + 2))
    Next i
End Sub

Sub PopulateTimeIndex(timeBlockIndex As Variant)
    ReDim timeBlockIndex(26)
    Worksheets("IndexValues").Activate
    For i = 0 To 26
        timeBlockIndex(i) = Worksheets("IndexValues").Range("H" & (i + 2))
    Next i
End Sub

Sub PopulateTimeNumberIndex(timeNumberIndex As Variant)
    ReDim timeNumberIndex(26)
    Worksheets("IndexValues").Activate
    For i = 0 To 26
        timeNumberIndex(i) = Worksheets("IndexValues").Range("J" & (i + 2))
    Next i
End Sub
