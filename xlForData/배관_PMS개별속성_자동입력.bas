Sub 배관개별속성입력()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 워크시트 정의
    Dim wb As Workbook, ws As Worksheet, pipWb As Workbook, pipWs As Worksheet
    Dim sizeWs As Worksheet
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("PMS")
    Set sizeWs = wb.Sheets("사이즈 변환")

    ' 사용자로부터 워크북 및 워크시트 선택
    Set pipWb = GetWorkbookFromUser
    If pipWb Is Nothing Then Exit Sub

    Dim pipWsNm As String
    pipWsNm = InputBox("워크시트 이름")
    On Error Resume Next
    Set pipWs = pipWb.Sheets(pipWsNm)
    If pipWs Is Nothing Then
        MsgBox "워크시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' PMS 열 인덱스 가져오기
    Dim pmsCols As Collection
    Set pmsCols = New Collection
    Dim headers As Variant: headers = Array("code", "MIN (float)", "MAX (float)", "OPERATING CONDITION TEMPERATURE", "GENERAL PWHT", "GENERAL BASE MATERIAL", "GENERAL C.A(mm)", "GENERAL MATERIAL", "GENERAL RATING", "GENERAL END CONNECTION TYPE", "GENERAL SCHEDULE", "GENERAL NON DESTRUCTIVE TEST RATE")
    Dim i As Integer
    For i = LBound(headers) To UBound(headers)
        pmsCols.Add FindColLetter(1, headers(i), ws)
    Next i

    ' PMS 데이터 배열 저장
    Dim pmsData As Variant
    pmsData = ws.Range(ws.Cells(2, 1), ws.Cells(ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, UBound(headers) + 1)).Value

    ' 개별 속성 데이터 저장
    Dim visibleCells As Range, tagData As Variant, tagList As Collection
    Set tagList = New Collection

    Set visibleCells = pipWs.Range(pipWs.Cells(2, pipWs.Range("속성 그룹 코드" & "1").Column), pipWs.Cells(pipWs.Cells(ws.Rows.Count, 1).End(xlUp).Row, pipWs.Range("속성 그룹 코드" & "1").Column)).SpecialCells(xlCellTypeVisible)
    For Each visibleCell In visibleCells
        If InStr(1, visibleCell.Value, "03") > 0 Then
            tagData = Array(pipWs.Cells(visibleCell.Row, pipWs.Range("개별속성8" & "1").Column).Value, GetFloatSize(pipWs.Cells(visibleCell.Row, pipWs.Range("개별속성9" & "1").Column).Value, sizeWs))
            tagList.Add tagData
        End If
    Next visibleCell

    ' 속성 데이터 입력
    Dim rowNum As Long
    Application.StatusBar = "작업을 시작합니다..."
    For i = 1 To tagList.Count
        For rowNum = LBound(pmsData) To UBound(pmsData)
            If tagList(i)(1) >= pmsData(rowNum, 2) And tagList(i)(1) <= pmsData(rowNum, 3) Then
                pipWs.Cells(i + 1, pipWs.Range("GENERAL PWHT" & "1").Column).Value = pmsData(rowNum, 4)
                pipWs.Cells(i + 1, pipWs.Range("GENERAL BASE MATERIAL" & "1").Column).Value = pmsData(rowNum, 5)
            End If
        Next rowNum
        If i Mod 10 = 0 Then Application.StatusBar = "진행 중: " & i & "/" & tagList.Count & " 완료..."
    Next i

    MsgBox "완료"
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub
