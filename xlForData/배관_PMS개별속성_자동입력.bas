Sub 배관개별속성입력()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 워크시트 정의
    Dim wb As Workbook, ws As Worksheet, pipWb As Workbook, pipWs As Worksheet
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("PMS")

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
    Dim pipColGroupCode As String, pipColLineClass As String, pipColLineSize As String
    pipColGroupCode = FindColLetter(1, "속성 그룹 코드", pipWs)
    pipColLineClass = FindColLetter(1, "개별속성8", pipWs)
    pipColLineSize = FindColLetter(1, "개별속성9", pipWs)

    Dim visibleCells As Range, tagData As Variant, tagList As Collection
    Set tagList = New Collection

    Set visibleCells = pipWs.Range(pipWs.Cells(2, pipWs.Range(pipColGroupCode & "1").Column), pipWs.Cells(pipWs.Cells(pipWs.Rows.Count, pipWs.Range(pipColGroupCode & "1").Column).End(xlUp).Row, pipWs.Range(pipColGroupCode & "1").Column)).SpecialCells(xlCellTypeVisible)

    For Each visibleCell In visibleCells
        If InStr(1, visibleCell.Value, "03") > 0 Then
            tagData = Array(pipWs.Cells(visibleCell.Row, pipWs.Range(pipColLineClass & "1").Column).Value, pipWs.Cells(visibleCell.Row, pipWs.Range(pipColLineSize & "1").Column).Value)
            tagList.Add tagData
        End If
    Next visibleCell

    ' 속성 데이터 입력
    Dim rowNum As Long
    Dim pmsColPwht As Integer, pmsColBaseMat As Integer
    pmsColPwht = FindColLetter(1, "GENERAL PWHT", pipWs)
    pmsColBaseMat = FindColLetter(1, "GENERAL BASE MATERIAL", pipWs)

    Application.StatusBar = "작업을 시작합니다..."
    For i = 1 To tagList.Count
        For rowNum = LBound(pmsData) To UBound(pmsData)
            If tagList(i)(1) >= pmsData(rowNum, 2) And tagList(i)(1) <= pmsData(rowNum, 3) Then
                pipWs.Cells(i + 1, pmsColPwht).Value = pmsData(rowNum, 4)
                pipWs.Cells(i + 1, pmsColBaseMat).Value = pmsData(rowNum, 5)
            End If
        Next rowNum
        If i Mod 10 = 0 Then Application.StatusBar = "진행 중: " & i & "/" & tagList.Count & " 완료..."
    Next i

    MsgBox "완료"
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub

Function GetWorkbookFromUser() As Workbook
    Dim wb As Workbook
    Dim filePath As String
    Dim fileName As String
    Dim isWorkbookOpen As Boolean

    ' 파일 선택 창 표시
    filePath = Application.GetOpenFilename("엑셀 파일(*.xls;*.xlsx;*.xlsb;*.xlsm), *.xls;*.xlsx;*.xlsb;*.xlsm", , "파일 선택", , False)

    ' 파일 선택 창에서 취소 버튼을 누른 경우
    If filePath = "False" Then
        MsgBox "취소", vbExclamation
        Set GetWorkbookFromUser = Nothing
        Exit Function
    End If

    ' 선택한 파일의 이름 가져오기
    fileName = Dir(filePath)

    ' 파일이 이미 열려 있는지 확인
    isWorkbookOpen = False
    For Each wb In Workbooks
        If wb.Name = fileName Then
            isWorkbookOpen = True
            Set GetWorkbookFromUser = wb
            Exit Function
        End If
    Next wb

    ' 파일이 열려 있지 않으면 열기
    If Not isWorkbookOpen Then
        Set wb = Workbooks.Open(filePath)
        Set GetWorkbookFromUser = wb
    End If

End Function

Function FindColLetter(hdr_row As Integer, search_value As Variant, Optional ws As Worksheet = Nothing) As String

    Dim search_rng As Range
    Dim found_cell As Range
    Dim col_letter As String

    ' 워크시트 변수를 설정. 기본값은 ActiveSheet
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If

    Set search_rng = ws.Rows(hdr_row)

    Set found_cell = search_rng.Find(What:=search_value, LookIn:=xlValues, LookAt:=xlWhole)

    If Not found_cell Is Nothing Then
        col_letter = Replace(found_cell.Cells.Address(False, False), hdr_row & "", "")
        FindColLetter = col_letter
    Else
        FindColLetter = "Value not found."
        MsgBox (search_value & " is not found")
    End If 

End Function
