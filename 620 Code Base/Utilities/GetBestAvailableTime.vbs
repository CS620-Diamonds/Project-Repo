Sub getBestAvailableTime(scheduleArray, timeBlockIndex, TotalScheduleRows, bestAvailableTime)
    Dim AvailableTimes(27) As String
    Dim availableTimesIndex As Integer
    availableTimesIndex = 0

    'find times already taken
    For i = 0 To 26
        For x = 1 To TotalScheduleRows
            If timeBlockIndex(i) <> scheduleArray(1, x - 1) Then
                AvailableTimes(availableTimesIndex) = timeBlockIndex(i)
                availableTimesIndex = availableTimesIndex + 1
                If availableTimesIndex = 27 Then
                    Exit For
                End If
            End If
        Next x
    Next i

    'find best remaining time
    bestAvailableTime = AvailableTimes(0)
End Sub
