Sub ImportData()
    On Error Resume Next
    Dim objExcel, oFSO, objImportedDataWorkbook, surveysFolder, objProfessorResponseWorkbook, ProfessorResponse, schedulerFile, FailCount, LoadCount
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    FailCount = 0
    LoadCount = 0

    'Set file directory containing survey files
    surveysFolder = "C:\Users\Jason.Perfetto.CORP\Documents\CS620\Surveys"

    'Set location of scheduler file
    schedulerFile = ActiveWorkbook.Name

    ' activate Scheduler file
    Set objSchedulerWorkbook = Workbooks(schedulerFile)
    objSchedulerWorkbook.Activate

    schedulerAC = objSchedulerWorkbook.Worksheets("Responses").Range("A1").Address
    Range(schedulerAC).Select

    'Check if folder has loadable surveys
    If Right(surveysFolder, 1) <> "\" Then
        surveysFolder = surveysFolder & "\"
    End If
    If Dir(surveysFolder & "*.*") = "" Then
        MsgBox "There are no surveys to load. Please put survey files in your 'Survey' directory"
        Exit Sub
    End If

        'create the excel object
    Set objExcel = CreateObject("Excel.Application")

    'view the excel program and file, set to false to hide the whole process
    objExcel.Visible = True

    'import responses
    For Each oFile In oFSO.GetFolder(surveysFolder).Files
        Err.Clear
        schedulerAC = ActiveCell.Address
        ActiveCell.Select
        If UCase(oFSO.GetExtensionName(oFile.Name)) = "XLS" Or UCase(oFSO.GetExtensionName(oFile.Name)) = "XLSX" Then
            If InStr(UCase(oFile.Name), "SURVEY") > 0 Then
                ProfessorResponse = oFile.Name
            End If
        End If

      ' open Professor Response file
        Set objProfResponseWorkbook = objExcel.Workbooks.Open(surveysFolder & "\" & ProfessorResponse, ignoreReadOnlyRecommended:=True)

      ' Select the range on Prof Survey you want to copy
        objProfResponseWorkbook.Worksheets("ProfessorAnswers").Activate
        objProfResponseWorkbook.Worksheets("ProfessorAnswers").Range("A1:AA1").Copy

      ' Paste it on Responses Sheet1, starting at A1
        objSchedulerWorkbook.Worksheets("Responses").Activate
        objSchedulerWorkbook.Worksheets("Responses").Range(schedulerAC).PasteSpecial xlPasteValues

      ' close the Survey workbook
        objProfResponseWorkbook.Close savechanges:=False

        'increment down to next row
        ActiveCell.Offset(1, 0).Select

        'catch/count errors
        If Err.Number <> 0 Then
            oFSO.MoveFile surveysFolder & "\" & ProfessorResponse, surveysFolder & "\failed\"
            FailCount = FailCount + 1
            WScript.Quit (Err.Number)
        End If

        oFSO.MoveFile surveysFolder & "\" & ProfessorResponse, surveysFolder & "\loaded\"
        LoadCount = LoadCount + 1
    Next

    ' Activate Responses Sheet1 so you can see it actually pasted the data
    objSchedulerWorkbook.Worksheets("Responses").Activate

    If ProfessorResponse <> "" Then
        If FailCount = 0 Then
            MsgBox LoadCount & " surveys have been loaded. " & FailCount & " surveys failed to load. Please close the blank instance of Excel and continue."
        End If
        If FailCount <> 0 Then
             MsgBox LoadCount & " surveys have been loaded. " & FailCount & " surveys failed to load. Please try to reload failed surveys."
        End If
    End If

End Sub