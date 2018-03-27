Sub getBestAvailableCourse(scheduleArray, courseIndex, TotalScheduleRows, bestAvailableCourse)
    Dim AvailableCourses(53) As String
    Dim availableCourseIndex As Integer
    availableCourseIndex = 0
    'find courses already taken
    For i = 0 To 53
        For x = 1 To TotalScheduleRows
            If courseIndex(i) <> scheduleArray(2, x - 1) Then
                AvailableCourses(availableCourseIndex) = courseIndex(i)
                availableCourseIndex = availableCourseIndex + 1
                If availableCourseIndex = 54 Then
                    Exit For
                End If
            End If
        Next x
    Next i

    'find best remaining time
    bestAvailableCourse = AvailableCourses(0)
End Sub
