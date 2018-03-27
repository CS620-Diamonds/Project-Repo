
Sub PopulateResponseArray(responseArray As Variant, TotalResponseRows)
    Worksheets("Responses").Activate
    TotalResponseRows = Worksheets("Responses").Rows(Rows.Count).End(xlUp).Row
    ReDim responseArray(6, TotalResponseRows)
    For i = 0 To UBound(responseArray, 2)
        responseArray(0, i) = Worksheets("Responses").Range("A1").Value  'faculty name
        responseArray(1, i) = Worksheets("Responses").Range("B1").Value  'full or part time INDEX
        responseArray(2, i) = Worksheets("Responses").Range("C1").Value  'course prefs
        responseArray(3, i) = Worksheets("Responses").Range("D1").Value  'time prefs
        responseArray(4, i) = Worksheets("Responses").Range("E1").Value 'back to back flag
        responseArray(5, i) = Worksheets("Responses").Range("F1").Value 'priority
    Next i
End Sub
