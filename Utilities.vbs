Sub printToSchedule(faculty, bestAvailableTime, bestAvailableCourse)
    Worksheets("Schedule").Activate
    For i = 0 To UBound(bestAvailableTime)
        If bestAvailableTime(i) <> "" And bestAvailableCourse(i) <> "" Then
            Worksheets("Schedule").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Value = faculty 'outputs Faculty Name attribute to empty row in A column
            Worksheets("Schedule").Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Value = bestAvailableTime(i)   'outputs Faculty Priority attribute to empty row in B column
            Worksheets("Schedule").Range("C" & Rows.Count).End(xlUp).Offset(1, 0).Value = bestAvailableCourse(i) 'outputs Faculty Course Pref varible to empty row in C column
            counter = i + 1
        End If
    Next i
    If counter = 0 Then
        Worksheets("Schedule").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Value = faculty
        Worksheets("Schedule").Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Value = "No Available Courses Match Preferences"
    End If
End Sub

Sub PopulateResponseArray(responseArray As Variant, TotalResponseRows)
    Worksheets("Responses").Activate
    TotalResponseRows = Worksheets("Responses").Rows(Rows.Count).End(xlUp).Row
    ReDim responseArray(6, TotalResponseRows)
    For i = 0 To TotalResponseRows
        responseArray(0, i) = Worksheets("Responses").Range("A" & i + 1).Value 'faculty name
        responseArray(1, i) = Worksheets("Responses").Range("B" & i + 1).Value  'full or part time INDEX
        responseArray(2, i) = Worksheets("Responses").Range("C" & i + 1).Value  'course prefs
        responseArray(3, i) = Worksheets("Responses").Range("D" & i + 1).Value  'time prefs
        responseArray(4, i) = Worksheets("Responses").Range("E" & i + 1).Value 'back to back flag
        responseArray(5, i) = Worksheets("Responses").Range("F" & i + 1).Value 'priority
    Next i
End Sub

Sub AddToArray(faculty, priorityArray As Variant, position)
    ReDim priorityArray(0 To position) As Variant
    priorityArray(position) = faculty
End Sub


Sub RefreshScheduleArray(scheduleArray As Variant, TotalScheduleRows)
    Worksheets("Schedule").Activate
    TotalScheduleRows = Worksheets("Schedule").Rows(Rows.Count).End(xlUp).Row
    ReDim scheduleArray(4, TotalScheduleRows) As Variant
    For i = 2 To TotalScheduleRows
            scheduleArray(1, i - 1) = Worksheets("Schedule").Cells(n, 1).Value 'writes facultyName to first attribute of each index
            scheduleArray(2, i - 1) = Worksheets("Schedule").Cells(n, 2).Value 'writes priority to second attribute of each index
            scheduleArray(3, i - 1) = Worksheets("Schedule").Cells(n, 3).Value 'writes course to third attribute of each index
            scheduleArray(4, i - 1) = Worksheets("Schedule").Cells(n, 4).Value 'writes timeblock to fourth attribute of each index
    Next i
End Sub

Sub isFullTime(responseArray As Variant, faculty, status, isFT, maxClasses)
    For i = 0 To UBound(responseArray, 2)
        If responseArray(0, i) = faculty Then
            status = responseArray(1, i)
            If status = "0" Then
                isFT = True
                maxClasses = 4
                Exit For
            End If
            If status <> "0" Then
                isFT = False
                maxClasses = 2
                Exit For
            End If
        End If
    Next i
End Sub

Sub getPriority(responseArray As Variant, faculty, priority)

    For i = 0 To UBound(responseArray)
        If responseArray(0, i) = faculty Then
            priority = responseArray(5, i)
            Exit For
        End If
    Next i

End Sub

Sub getCoursePreferences(responseArray As Variant, faculty, courses)
    ReDim courses(0 To 10) As Variant
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

Sub getTimePrefs(responseArray As Variant, faculty, times)
    ReDim times(0 To 26) As Variant
    Dim time1 As String
    Dim time2 As String
    Dim time3 As String
    Dim time4 As String
    Dim time5 As String
    Dim time6 As String
    Dim time7 As String
    Dim time8 As String
    Dim time9 As String
    Dim time10 As String
    Dim time11 As String
    Dim time12 As String
    Dim time13 As String
    Dim time14 As String
    Dim time15 As String
    Dim time16 As String
    Dim time17 As String
    Dim time18 As String
    Dim time19 As String
    Dim time20 As String
    Dim time21 As String
    Dim time22 As String
    Dim time23 As String
    Dim time24 As String
    Dim time25 As String
    Dim time26 As String
    Dim time27 As String

For i = 0 To UBound(responseArray, 2)
        If responseArray(0, i) = faculty Then
            times = Split(responseArray(3, i), "[]")
            time1 = times(0)
            time2 = times(1)
            time3 = times(2)
            time4 = times(3)
            time5 = times(4)
            time6 = times(5)
            time7 = times(6)
            time8 = times(7)
            time9 = times(8)
            time10 = times(9)
            time11 = times(10)
            time12 = times(11)
            time13 = times(12)
            time14 = times(13)
            time15 = times(14)
            time16 = times(15)
            time17 = times(16)
            time18 = times(17)
            time19 = times(18)
            time20 = times(19)
            time21 = times(20)
            time22 = times(21)
            time23 = times(22)
            time24 = times(23)
            time25 = times(24)
            time26 = times(25)
            time27 = times(26)
            Exit For
        End If
    Next i
End Sub

Sub getBestAvailableTime(scheduleArray, timeNumberIndex, TotalScheduleRows, times, maxClasses, bestAvailableTime)
    Dim AvailableTimes(26) As String
    Dim availableTimesIndex As Integer
    ReDim bestAvailableTime(0 To maxClasses - 1)
    Dim maxCoursesPerTime As Integer
    maxCoursesPerTime = 2

'    'find times already taken
    For i = 0 To 26
    availableTimesIndex = 0
        AvailableTimes(i) = timeNumberIndex(i)
        availableTimesIndex = availableTimesIndex + 1
        If availableTimesIndex = 27 Then
            Exit For
        End If
    Next i
    'count instances of times
    timeLoad = 0
    For i = 0 To UBound(AvailableTimes)
        timeBlockMatch = AvailableTimes(i)
        For x = 0 To UBound(AvailableTimes)
            If IsInArray(timeBlockMatch, AvailableTimes) Then
                timeLoad = timeLoad + 1
            End If
        Next x
    Next i


     'find best remaining time
    bestAvailableTime(0) = ""
    For x = 0 To UBound(bestAvailableTime)
        If times(x) = "-99" Then
            Exit For
        End If
        timeBlockMatch = AvailableTimes(i)
        For n = 0 To UBound(AvailableTimes)
            If IsInArray(timeBlockMatch, AvailableTimes) Then
                timeLoad = timeLoad + 1
            End If
            If times(x) = AvailableTimes(n) Then
                bestAvailableTime(x) = AvailableTimes(n)
            End If
        Next n
    Next x
End Sub

Sub getBestAvailableCourse(scheduleArray, courseIndex, TotalScheduleRows, courses, maxClasses, bestAvailableCourse)
    Dim AvailableCourses(53) As String
    Dim takenCourses(53) As String
    Dim availableCourseIndex As Integer
    Dim profCourse As String
    ReDim bestAvailableCourse(0 To maxClasses - 1)
    Worksheets("Schedule").Activate
    TotalScheduleRows = Worksheets("Schedule").Rows(Rows.Count).End(xlUp).Row
    Worksheets("Schedule").Range("C2").Select

    For i = 0 To TotalScheduleRows
        Worksheets("Schedule").Range("C" & i + 2).Select
        courseLooper = ActiveCell.Address
        courseToMatch = ActiveCell.Value
        For x = 0 To 53
            If courseIndex(x) = courseToMatch Then
                takenCourses(i) = courseIndex(x)
                availableCourseIndex = availableCourseIndex + 1
                If availableCourseIndex = 54 Then
                    Exit For
                End If
            End If
        Next x
        ActiveCell.Offset(1, 0).Select
    Next i
    'find best remaining course
    bestAvailableCourse(0) = ""
    availableCourseIndex = 0
    For x = 0 To UBound(bestAvailableCourse)
        For n = 0 To UBound(courses)
            profCourse = courses(n)
            If profCourse = "-99" Then
                Exit For
            End If
            If availableCourseIndex = maxClasses Then
                Exit For
            End If
            If IsInArray(profCourse, takenCourses) = False Then
                bestAvailableCourse(availableCourseIndex) = profCourse
                availableCourseIndex = availableCourseIndex + 1
            End If
        Next n
        If availableCourseIndex = maxClasses Then
            Exit For
        End If
        If profCourse = "-99" Then
            Exit For
        End If
    Next x

End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = UBound(Filter(arr, stringToBeFound)) > -1
End Function


Sub NewScheduleSheet()

    Dim ws As Worksheet  'declariation of ws
    Dim found As Boolean  'declariation of found

    found = False  'sets variable found

    For Each ws In ThisWorkbook.Sheets  'finds if there is a worksheet called Schedule if exists it exits
        If ws.Name = "Schedule" Then
            found = True
            Exit For
        End If

    Next
    If Not found Then  'If worksheet schedule is not found it is created and columns are added.
        Sheets.Add.Name = "Schedule"
        Worksheets("Schedule").Range("A1").Value = "Faculty"
        Worksheets("Schedule").Range("B1").Value = "TimeBlock"
        Worksheets("Schedule").Range("C1").Value = "Course"
    End If
    Worksheets("Responses").Activate
End Sub

Sub hasPrerequisite(course, hasPrerequisites)
    Dim colLoopAddress
    Dim rowLoopAddress
    position = InStr(course, ("|"))
    course = Left(course, position - 1)
    Worksheets("CISPrerequisites").Activate
    Worksheets("CISPrerequisites").Range("A1").Select
    colLoopAddress = ActiveCell.Address

    'find courses already taken
    For i = 2 To 16
        If ActiveCell.Value = course Then
            rowLoopAddress = ActiveCell.Address
            Exit For
        End If
        colLoopAddress = ActiveCell.Offset(1, 0).Address
        Range(colLoopAddress).Select
    Next i
        For x = 2 To 16
            Select Case ActiveCell.Value
                Case 1
                    hasPrerequisites = True
                Case 2
                    hasPrerequisites = True
            End Select
            ActiveCell.Offset(0, 1).Select
            rowLoopAddress = ActiveCell.Address
    Next x
End Sub

Sub LookupTime(bestAvailableTime, timeBlockIndex)
    For i = 0 To UBound(bestAvailableTime)
        If timeIndex <> 0 Then
            timeIndex = CInt(bestAvailableTime(i))
            bestAvailableTime(i) = timeBlockIndex(timeIndex - 1)
        End If
    Next i
End Sub
