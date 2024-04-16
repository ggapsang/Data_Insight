Sub SwapRanges()
  Dim range1 As Range
  Dim range2 As Range
  Dim tempValue As Variant
  Dim i As Integer
  Dim inputRange As Range ' 여기를 Range 타입으로 변경
    
  ' 사용자가 선택한 첫 번째 범위
  Set range1 = Selection
    
  ' 두 번째 범위 입력 요청
  On Error Resume Next ' 사용자가 취소할 경우 오류 방지
  Set inputRange = Application.InputBox("두 번째 범위 입력:", Type:=8)
  On Error GoTo 0 ' 오류 처리를 기본으로 돌림
    
  ' 사용자가 취소한 경우 체크
  If inputRange Is Nothing Then
      MsgBox "작업 취소."
      Exit Sub
  End If
    
  ' 입력받은 두 번째 범위 설정
  Set range2 = inputRange
  
  ' 두 범위의 크기 비교
  If range1.Rows.Count <> range2.Rows.Count Or range1.Columns.Count <> range2.Columns.Count Then
      MsgBox "범위의 크기가 일치하지 않음."
      Exit Sub
  End If
    
  ' 값 교환
  For i = 1 To range1.Rows.Count
      For j = 1 To range1.Columns.Count ' 열도 동적으로 처리
          tempValue = range1.Cells(i, j).value
          range1.Cells(i, j).value = range2.Cells(i, j).value
          range2.Cells(i, j).value = tempValue
      Next j
  Next i

End Sub
