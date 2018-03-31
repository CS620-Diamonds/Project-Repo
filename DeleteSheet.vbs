
 Sub DeleteSchedule() 'deletes sheet called "Schedule" if there is one - resets active cell to Data Import A2
    Dim ws As Worksheet  'declariation of ws
    Dim found As Boolean  'declariation of found

    Application.DisplayAlerts = False  'Turns off warnings and confirmations

    found = False  'sets variable found

    For Each ws In ThisWorkbook.Sheets  'finds if there is a worksheet called Schedule and deletes it
        If ws.Name = "Schedule" Then
            found = True
            Sheets("Schedule").Select
            ActiveWindow.SelectedSheets.Delete
            Exit For
        End If

    Next
    If Not found Then       'if worksheet schedule is not found exit
        MsgBox "There is no schedule to delete"
    End If

    Sheets("Responses").Select  'Select A1 on Repsonses Sheet
    Worksheets("Responses").Range("A1").Select
    Selection.Style = "Normal"

    Application.DisplayAlerts = True  'Turns on warnings and confirmations
End Sub
