Function RCOUNTA(rng As Range)

    Dim cell As Range
    Dim count As Long
    
    count = 0
    
    For Each cell In rng

        If cell.value <> "" Then
        count = count + 1
        End If

    Next cell
    
    RCOUNTA = count

End Function
