Sub printToSchedule(faculty, time, course)
    Worksheets("Schedule").Activate
    Worksheets("Schedule").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Value = faculty 'outputs Faculty Name attribute to empty row in A column
    Worksheets("Schedule").Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Value = time   'outputs Faculty Priority attribute to empty row in B column
    Worksheets("Schedule").Range("C" & Rows.Count).End(xlUp).Offset(1, 0).Value = course 'outputs Faculty Course Pref varible to empty row in C column
End Sub
