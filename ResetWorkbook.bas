Attribute VB_Name = "ResetWorkbook"
Sub ResetWorkbook()
Attribute ResetWorkbook.VB_Description = "This macro resets the Loop workbook"
Attribute ResetWorkbook.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ResetWorkBook Macro
' This macro resets the Loop workbook
' This is only for quickly resetting the page logic
'

'
Dim ws As Worksheet  'declariation of ws
Dim found As Boolean  'declariation of found

Application.DisplayAlerts = False  'Turns off warnings and confirmations

found = False  'sets variable found

For Each ws In ThisWorkbook.Sheets  'finds if there is a worksheet called Schedule and deletes it.
    If ws.Name = "Schedule" Then
        found = True
        Sheets("Schedule").Select
        ActiveWindow.SelectedSheets.Delete
        Exit For
    End If

Next
If Not found Then  'if worksheet schedule is not found exit
End If
    
    Sheets("DataImport").Select  'sets formatting on DataImport sheet and sets it active
    Columns("A:C").Select
    Selection.Style = "Normal"

Application.DisplayAlerts = True  'Turns on warnings and confirmations

End Sub
