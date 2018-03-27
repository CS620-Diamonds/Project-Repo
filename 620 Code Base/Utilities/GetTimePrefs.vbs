Sub getTimePrefs(responseArray As Variant, faculty)
    Dim times() As String
    Dim time1 As String
    Dim time2 As String
    Dim time3 As String
    Dim time4 As String
    Dim time5 As String
    Dim time6 As String
    Dim time7 As String
    Dim time8 As String
    Dim time9 As String
    Dim time10 As String
    Dim time11 As String
    Dim time12 As String
    Dim time13 As String
    Dim time14 As String
    Dim time15 As String
    Dim time16 As String
    Dim time17 As String
    Dim time18 As String
    Dim time19 As String
    Dim time20 As String
    Dim time21 As String
    Dim time22 As String
    Dim time23 As String
    Dim time24 As String
    Dim time25 As String
    Dim time26 As String
    Dim time27 As String

For i = 0 To UBound(responseArray, 2)
        If responseArray(0, i) = faculty Then
            times() = Split(responseArray(3, i), "[]")
            time1 = times(0)
            time2 = times(1)
            time3 = times(2)
            time4 = times(3)
            time5 = times(4)
            time6 = times(5)
            time7 = times(6)
            time8 = times(7)
            time9 = times(8)
            time10 = times(9)
            time11 = times(10)
            time12 = times(11)
            time13 = times(12)
            time14 = times(13)
            time15 = times(14)
            time16 = times(15)
            time17 = times(16)
            time18 = times(17)
            time19 = times(18)
            time20 = times(19)
            time21 = times(20)
            time22 = times(21)
            time23 = times(22)
            time24 = times(23)
            time25 = times(24)
            time26 = times(25)
            time27 = times(26)
            Exit For
        End If
    Next i
End Sub
