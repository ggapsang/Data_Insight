''' 유저 폼 형성 가이드
''' 1. 사용자 정의 폼 생성
''' 2. 폼에 컨트롤 추가
''' 3. 콤보박스 해더 추가
''' 4. 콤보박스를 선택하면 해당 칼럼의 고유값을 리스트박스에 표시
''' 리스트박스에서 값들을 선택하고 확인 버튼 클릭 시 선택된 값들에 해당하는 행을 새로운 워크북에 복사 후 저장

'' ComobBox : cmbColumns
'' ListBox : lstUniqueValues, 'MulitSelect' 속성을 '1 - fmMultiSelectMulti'로 설정
'' Button : btnConfirm, 캡션 : 확인


''' 유저 폼 초기화
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim i As Long
    For i = 1 To lastCol
        cmbColumns.AddItem ws.Cells(1, i).Value
    Next i
End Sub

''' 콤보박스 변경 이벤트
Private Sub cmbColumns_Change()
    lstUniqueValues.Clear

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim colIndex As Long
    colIndex = Application.Match(cmbColumns.Value, ws.Rows(1), 0)

    Dim uniqueValues As Collection
    Set uniqueValues = New Collection

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

    Dim i As Long
    On Error Resume Next
    For i = 2 To lastRow
        uniqueValues.Add ws.Cells(i, colIndex).Value, CStr(ws.Cells(i, colIndex).Value)
    Next i
    On Error GoTo 0

    Dim val As Variant
    For Each val In uniqueValues
        lstUniqueValues.AddItem val
    Next val
End Sub

''' 확인 버튼 클릭 이벤트
Private Sub btnConfirm_Click()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim colIndex As Long
    colIndex = Application.Match(cmbColumns.Value, ws.Rows(1), 0)

    Dim selectedValues As Collection
    Set selectedValues = New Collection

    Dim i As Long
    For i = 0 To lstUniqueValues.ListCount - 1
        If lstUniqueValues.Selected(i) Then
            selectedValues.Add lstUniqueValues.List(i)
        End If
    Next i

    ' 선택된 값이 없는 경우 메시지 표시 후 종료
    If selectedValues.Count = 0 Then
        MsgBox "선택된 값이 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' 최적화 시작
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    

    Dim criteria As String
    criteria = ""
    For i = 1 To selectedValues.Count
        criteria = criteria & selectedValues(i) & ","
    Next i
    criteria = Left(criteria, Len(criteria) - 1) ' 마지막 쉼표 제거

    ws.Range("A1").AutoFilter Field:=colIndex, Criteria1:=Split(criteria, ","), Operator:=xlFilterValues

    Dim newWb As Workbook
    Set newWb = Workbooks.Add

    Dim newWs As Worksheet
    Set newWs = newWb.Sheets(1)

    ' 필터링된 데이터를 복사
    ws.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=newWs.Range("A1")

    ' 필터 해제
    ws.AutoFilterMode = False

    Dim fileName As String
    fileName = Application.GetSaveAsFilename("", "Excel Files (*.xlsx), *.xlsx")

    If fileName <> "False" Then
        newWb.SaveAs fileName
        newWb.Close
    Else
        MsgBox "파일 저장이 취소되었습니다.", vbInformation
    End If

    ' 최적화 종료
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    ' 유저폼 닫기
    Unload Me
    
End Sub


Private Function IsInCollection(col As Collection, val As Variant) As Boolean
    Dim item As Variant
    IsInCollection = False
    For Each item In col
        If item = val Then
            IsInCollection = True
            Exit For
        End If
    Next item
End Function



