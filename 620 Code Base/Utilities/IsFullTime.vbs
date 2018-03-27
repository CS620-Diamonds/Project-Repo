Sub isFullTime(responseArray As Variant, faculty, status, isFT, maxClasses)
    For i = 0 To UBound(responseArray, 2)
        If responseArray(0, i) = faculty Then
            status = responseArray(1, i)
            If status = "1" Then
                isFT = True
                maxClasses = 4
                Exit For
            End If
            If status <> "1" Then
                isFT = False
                maxClasses = 2
                Exit For
            End If
        End If
    Next i
End Sub
