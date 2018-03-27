Sub getPriority(responseArray As Variant, faculty, priority)

    For i = 0 To UBound(responseArray)
        If responseArray(0, i) = faculty Then
            priority = responseArray(5, i)
            Exit For
        End If
    Next i

End Sub
