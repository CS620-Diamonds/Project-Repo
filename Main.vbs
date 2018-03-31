Sub Main()
    Call Initialize.DeclareGlobals
    'Call Initialize.Initialize  'sets globals and populates index arrays
    Call Utilities.PopulateResponseArray(responseArray, TotalResponseRows) 'stores response data in an array
    Call Utilities.RefreshScheduleArray(scheduleArray, TotalScheduleRows)  'refreshes schedule data in an array
    Call Initialize.PopulateTimeIndex(timeBlockIndex)
    Call Initialize.PopulateB2BFlag(back2backFlag)
    Call Initialize.PopulateCourseIndex(courseIndex)
    Call Initialize.PopulateDaysOffFlag(daysOffFlag)
    Call Initialize.PopulateDaysOffIndex(daysOffIndex)
    Call Initialize.PopulateFacultyIndex(facultyIndex)
    Call Initialize.PopulateGradUndergradIndex(gradUndergradIndex)
    Call Initialize.PopulateStatusIndex(statusIndex)
    Call Initialize.PopulateTimeNumberIndex(timeNumberIndex)


'Schedule Classes
        Worksheets("Responses").Activate
        Worksheets("Responses").Range("A1").Select
        responseLoopAddress = ActiveCell.Address
        For i = 1 To TotalResponseRows
            Worksheets("Responses").Activate
            Range("A" & i).Select
            faculty = ActiveCell.Value
            Call Utilities.isFullTime(responseArray, faculty, status, isFT, maxClasses)
            Call Utilities.getTimePrefs(responseArray, faculty, times)
            Call Utilities.getCoursePreferences(responseArray, faculty, courses)

            If isFT Then
                maxClasses = 4
            Else
                maxClasses = 2
            End If

            Call Utilities.getBestAvailableTime(scheduleArray, timeNumberIndex, TotalScheduleRows, times, maxClasses, bestAvailableTime)
            Call Utilities.getBestAvailableCourse(scheduleArray, courseIndex, TotalScheduleRows, courses, maxClasses, bestAvailableCourse)
            Call Utilities.LookupTime(bestAvailableTime, timeBlockIndex)
            Call Utilities.printToSchedule(faculty, bestAvailableTime, bestAvailableCourse)
            ActiveCell.Offset(1, 0).Select
            responseLoopAddress = ActiveCell.Address
        Next i
End Sub
