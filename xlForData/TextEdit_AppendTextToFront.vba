Sub AppendTextToFront()

  Dim cell As Range
  Dim AppendText As String
    
  AppendText = InputBox("구분자 설정 : ")
  Application.DisplayAlerts = False

  For Each cell In Selection.Cells
  
      Dim value As Variant
      value = cell.value
      cell.value = AppendText & value
  
  Next cell
    
End Sub
