Sub DeclareGlobals() 'Global Variable Instantiation
        Dim maxPerBlock As Integer
        maxPerBlock = 3

        Dim isFT As Boolean
        Dim priority As Integer 'faculty priority
        Dim status As String 'faculty status
        Dim maxClasses As Integer 'faculty max classes
        Dim bestAvailableTime As String
        Dim bestAvailableCourse As String

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
        Dim scheduleArray(9, 1) As Variant 'Instantiate schedule array - each index is a Faculty Object with 7 attributes
        Dim scheduleFaculty(9, 1) As Variant 'Instantiate scheduleFaculty Object, from each index in ^
    'response
        Dim responseArray(9, 1) As Variant 'Instantiate response array - each index is a Faculty Object with 7 attributes
        Dim responseFaculty(9, 1) As Variant 'Instantiate responseFaculty Object, from each index in ^

    'Array dictionary:
    '1 : Name : String
    '2 : Status : String
    '3 : Grad/Undergrad : String
    '4 : Course Prefs (n = 10) : String
    '5 : Days Off Needed (n = 5) : String
    '6 : Times unavailable (n = 10) : String
    '7 : Time Preferences (n = 27) : String
    '8 : Back2Back (bool) : Boolean
    '9 : Priority (Integer)

    'Index Arrays
        Dim timeBlockIndex(26) As Variant
        Dim courseIndex(53) As Variant
        Dim professorIndex(32, 2) As Variant
        Dim statusIndex(3, 2) As Variant
        Dim gradUndergradIndex(2, 3) As Variant
        Dim daysOffFlag(2, 2) As Variant
        Dim daysOffIndex(5, 2) As Variant
        Dim back2backFlag(2, 2) As Variant

    'Split Values
        Dim i As Integer 'loop index
        Dim counter As Integer 'coursePreference Counter for assignment
        Dim strACAddress As Variant
        Dim strFaculty As String
        Dim timeCell

    'Get Sheet Lengths
        TotalResponseRows = Worksheets("Responses").Rows(Rows.Count).End(xlUp).Row
        TotalDataImportRows = Worksheets("DataImport").Rows(Rows.Count).End(xlUp).Row
        'TotalScheduleRows = Worksheets("Schedule").Rows(Rows.Count).End(xlUp).Row

End Sub
