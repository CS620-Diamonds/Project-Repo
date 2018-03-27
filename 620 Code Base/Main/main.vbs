Sub Main()
    Call Initialize.DeclareGlobals
    Call Initialize.Initialize  'sets globals and populates index arrays
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

   'for each response:
    For i = 1 To TotalResponseRows
        faculty = responseArray(0, i)
        Call Utilities.getPriority(responseArray, faculty, priority)  'get priority of faculty
        If responseArray(5, i) = priority Then
            Call Utilities.isFullTime(responseArray, faculty, status, isFT, maxClasses)
            Call Utilities.getTimePrefs(responseArray, faculty)
            Call Utilities.getCoursePreferences(responseArray, faculty)
            Call Utilities.getBestAvailableTime(scheduleArray, timeBlockIndex, TotalScheduleRows, bestAvailableTime)
            Call Utilities.getBestAvailableCourse(scheduleArray, courseIndex, TotalScheduleRows, bestAvailableCourse)
            Call Utilities.printToSchedule(faculty, bestAvailableTime, bestAvailableCourse)
        End If
   Next i
End Sub
