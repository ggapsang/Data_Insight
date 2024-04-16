Sub NormalizeCellContents()
    Dim replaceWith As String
    replaceWith = InputBox("구분자 설정 : ", "Input Required", ", ")
    ReplaceNewLines replaceWith
    UnmergeAndClean replaceWith
End Sub

Sub ReplaceNewLines(replaceWith As String)
    Dim cell As Range
    Dim value As Variant

    For Each cell In Selection.Cells
        value = cell.value
        If InStr(1, value, vbLf) > 0 Then
            value = Replace(value, vbLf, replaceWith)
            cell.value = value
        ElseIf InStr(1, value, vbCr) > 0 Then
            value = Replace(value, vbCr, replaceWith)
            cell.value = value
        End If
    Next cell
End Sub

Sub UnmergeAndClean(replaceWith As String)
    Dim cell As Range
    Dim value As Variant

    For Each cell In Selection.Cells
        If cell.MergeCells Then
            Dim mergeRange As Range
            Set mergeRange = cell.MergeArea
            cell.UnMerge
            value = cell.value
            value = Replace(value, vbCrLf, replaceWith)
            mergeRange.value = value
        End If
    Next cell
End Sub
