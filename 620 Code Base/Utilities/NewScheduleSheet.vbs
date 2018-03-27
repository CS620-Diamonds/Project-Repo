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
