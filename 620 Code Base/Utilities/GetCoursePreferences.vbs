Sub getCoursePreferences(responseArray As Variant, faculty)
    Dim courses() As String
    Dim course1 As String
    Dim course2 As String
    Dim course3 As String
    Dim course4 As String
    Dim course5 As String
    Dim course6 As String
    Dim course7 As String
    Dim course8 As String
    Dim course9 As String
    Dim course10 As String

    For i = 0 To UBound(responseArray)
        If responseArray(0, i) = faculty Then
            courses = Split(responseArray(2, i), "{}")
            course1 = courses(0)
            course2 = courses(1)
            course3 = courses(2)
            course4 = courses(3)
            course5 = courses(4)
            course6 = courses(5)
            course7 = courses(6)
            course8 = courses(7)
            course9 = courses(8)
            course10 = courses(9)
            Exit For
        End If
    Next i
End Sub
