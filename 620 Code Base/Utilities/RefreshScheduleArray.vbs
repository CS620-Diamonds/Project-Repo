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
