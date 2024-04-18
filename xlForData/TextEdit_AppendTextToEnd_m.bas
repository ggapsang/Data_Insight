Sub AppendTextToEnd()

    Dim header_row As Integer
    Dim cardNo_col As Integer
    Dim cardNo_col_chr As String
    Dim cell As Range
    
    Dim isSkipEmpty As Boolean
    Dim replaceWith As String
    
    Dim value As Variant
    
    Dim skipFlas As Boolean
    
    isSkipEmpty = Application.InputBox(prompt:="빈 셀 생략(true=1, false=0) : ", Type:=4)
    
    replaceWith = InputBox("뒤에 붙이기 :")
    
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each cell In Selection.Cells
    
        If isSkipEmpty = False Then
            
            value = cell.value
            cell.value = value & replaceWith
        
        Else

            value = cell.value
            skipFlag = value = ""
            
            If skipFlag = False Then
                cell.value = value & replaceWith
            End If

        End If
            
    Next cell

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True


End Sub
