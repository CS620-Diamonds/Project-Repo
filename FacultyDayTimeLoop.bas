Attribute VB_Name = "FacultyDayTimeLoop"

Sub FacultyDayTimeLoop()
'This is a simple nested, nested loop. GLM

'Checks for WorkSheet named Schedule. If missing it creates it.
Call NewScheduleSheet

'DataImport Section
Worksheets("DataImport").Activate
Worksheets("DataImport").Range(Cells(2, 1), ActiveCell).Select  'Sets Active Cell to CalcSheet.A2

Do
    ActiveCell.Font.Italic = True  'test for processing
    strACAddress = ActiveCell.Address  'records active cell address. Increments down one row each pass.
    strFaculty = ActiveCell.Value  'sets active cell as strFaculty variable
    Range("B2").Select   'selects first day entry B2
            Do   'day loop iteration
                strACLoopAddress = ActiveCell.Address  'records active Loop cell address. Increments down one row each pass.
                ActiveCell.Font.Bold = True  'test for processing
                strDay = ActiveCell.Value  'sets active cell as strDay variable
                Range("C2").Select  'selects first time entry C2
                    
                    Do
                        ActiveCell.Font.Bold = True  'test for processing
                        ActiveCell.Font.Italic = True  'test for processing
                        strTime = ActiveCell.Value  'sets active cell as strTime variable
                        Worksheets("Schedule").Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Value = strFaculty  'outputs strFaculty varible to empty row in A column
                        Worksheets("Schedule").Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Value = strDay  'outputs strDay varible to empty row in B column
                        Worksheets("Schedule").Range("C" & Rows.Count).End(xlUp).Offset(1, 0).Value = strTime  'outputs strTime varible to empty row in C column
                        ActiveCell.Offset(1, 0).Select  'Increments the active cell down one row
                    Loop Until ActiveCell.Value = ""  'Loops Time comlumn until it finds a blank cell then exits loop
                
                Range(strACLoopAddress).Select  'resets active cell to active Loop cell address - Day
                ActiveCell.Offset(1, 0).Select  'Increments the active cell down one row
            Loop Until ActiveCell.Value = ""  'Loops Day comlumn until it finds a blank cell then exits loop
    
    Range(strACAddress).Select  'resets active cell to active cell address - Faculty
    ActiveCell.Offset(1, 0).Select  'Increments the active cell down one row
Loop Until ActiveCell.Value = ""  'Loops Faculty comlumn until it finds a blank cell then exits loop

Worksheets("Schedule").Select  'Sets Schedule as the active sheet

End Sub

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
    Worksheets("Schedule").Range("B1").Value = "Day"
    Worksheets("Schedule").Range("C1").Value = "Time"
End If

End Sub
