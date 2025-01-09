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
    Dim pmsColCode As String 'pms 시트 code 열
    Dim pmsColMin As String 'pms 시트 MIN (float) 열
    Dim pmsColMax As String 'pms 시트 MAX (float) 열
    Dim pmsColOperCond As String 'pms 시트 Operating Condition 열
    Dim pmsColPwht As String 'pms 시트 PHWT 열
    Dim pmsColBaseMat As String 'pms 시트 BASE MATERIAL 열
    Dim pmsColCa As String 'pms 시트 C.A 열
    Dim pmsColRating As String 'pms 시트 RATING 열
    Dim pmsColMat As String 'pms 시트 GENERAL MATERIAL 열
    Dim pmsColConnectType As String 'pms 시트 END TYPE CONNECTION 열
    Dim pmsColSch As String 'pms 시트 GENERAL SCHEDULE 열
    Dim pmsColNdt As String 'pms 시트 GENERAL NON DESTRUCTIVE TEST RATE 열

    pmsColCode = FindColLetter(1, "code", ws)
    pmsColMin = FindColLetter(1, "MIN (float)", ws)
    pmsColMax = FindColLetter(1, "MAX (float)", ws)
    pmsColOperCond = FindColLetter(1, "OPERATING CONDITION TEMPERATURE", ws)
    pmsColPwht = FindColLetter(1, "GENERAL PWHT", ws)
    pmsColBaseMat = FindColLetter(1, "GENERAL BASE MATERIAL", ws)
    pmsColCa = FindColLetter(1, "GENERAL C.A(mm)", ws)
    pmsColRating = FindColLetter(1, "GENERAL RATING", ws)
    pmsColMat = FindColLetter(1, "GENERAL MATERIAL", ws)
    pmsColConnectType = FindColLetter(1, "GENERAL END CONNECTION TYPE", ws)
    pmsColSch = FindColLetter(1, "GENERAL SCHEDULE", ws)
    pmsColNdt = FindColLetter(1, "GENERAL NON DESTRUCTIVE TEST RATE", ws)

    ' PMS 데이터 배열 저장
    Dim pmsData() As Variant
    Dim lastRow As Long, lastCol As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    pmsData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value

    ' 개별 속성 데이터 저장
    Dim pipColTagNo As String '설비번호
    Dim pipColType As String '공정별 분류 코드 : 10029001
    Dim pipColGroupCode As String '속성 그룹 코드
    Dim pipColLineClass As String '개별속성1 : GENERAL LINE CLASS
    Dim pipColLineSize As String ' 개별속성2 : GENERAL LINE SIZE
    Dim pipColPwht As String '개별속성16 : GENERAL PWHT
    Dim pipColBaseMat As String '개별속성17 : GENERAL BASE MATERIAL
    Dim pipColCA As String '개별속성18 : KILLED CS
    Dim pipColRating As String '개별속성21 : GENERAL RATING
    Dim pipColMaterial As String '개별속성22 : GENERAL MATERIAL
    Dim pipColConType As String '개별속성24 : GENERAL END CONNECTION TYPE
    Dim pipColSch As String '개별속성25 : GENERAL SCHEDULE
    Dim pipColNDE As String '개별속성26 : GENERAL NON DESTRUCTIVE TEST RATE

    pipColTagNo = FindColLetter(1, "설비번호", pipWs) '설비번호
    pipColType = FindColLetter(1, "공정별 분류 코드", pipWs) '공정별 분류 코드 : 10029001
    pipColGroupCode = FindColLetter(1, "속성 그룹 코드", pipWs) '속성 그룹 코드
    pipColLineClass = FindColLetter(1, "개별속성8", pipWs) '개별속성1 : GENERAL LINE CLASS
    pipColLineSize = FindColLetter(1, "개별속성9", pipWs) ' 개별속성2 : GENERAL LINE SIZE
    pipColPwht = FindColLetter(1, "개별속성19", pipWs) '개별속성17 : GENERAL PWHT
    pipColBaseMat = FindColLetter(1, "개별속성20", pipWs) '개별속성18 : GENERAL BASE MATERIAL
    pipColCA = FindColLetter(1, "개별속성21", pipWs) '개별속성19 : KILLED CA
    pipColRating = FindColLetter(1, "개별속성22", pipWs) '개별속성21 : GENERAL RATING
    pipColMaterial = FindColLetter(1, "개별속성23", pipWs) '개별속성22 : GENERAL MATERIAL
    pipColConType = FindColLetter(1, "개별속성25", pipWs) '개별속성24 : GENERAL END CONNECTION TYPE
    pipColSch = FindColLetter(1, "개별속성26", pipWs) '개별속성25 : GENERAL SCHEDULE
    pipColNDE = FindColLetter(1, "개별속성27", pipWs) '개별속성26 : GENERAL NON DESTRUCTIVE TEST RATE

    Dim visibleCells As Range, tagData As Variant, tagList As Collection
    Set tagList = New Collection

    Set visibleCells = pipWs.Range(pipWs.Cells(2, pipWs.Range(pipColGroupCode & "1").Column), pipWs.Cells(pipWs.Cells(pipWs.Rows.Count, pipWs.Range(pipColGroupCode & "1").Column).End(xlUp).Row, pipWs.Range(pipColGroupCode & "1").Column)).SpecialCells(xlCellTypeVisible)

    For Each visibleCell In visibleCells
        If InStr(1, visibleCell.Value, "03") > 0 Then
            tagData = Array(pipWs.Cells(visibleCell.Row, pipWs.Range(pipColLineClass & "1").Column).Value, pipWs.Cells(visibleCell.Row, pipWs.Range(pipColLineSize & "1").Column).Value)
            tagList.Add tagData
        End If
    Next visibleCell

    Dim i As Long, rowNum As Long

    Application.StatusBar = "작업을 시작합니다..."
    For i = 1 To tagList.Count
        For rowNum = 2 To lastRow
            If tagList(i)(1) = pmsData(rowNum, ws.Range(pmsColCode & "1").Column) Then
                pipWs.Cells(i + 1, pipWs.Range(pipColPwht & "1").Column).Value = pmsData(rowNum, ws.Range(pmsColPwht & "1").Column)
                pipWs.Cells(i + 1, pipWs.Range(pipColBaseMat & "1").Column).Value = pmsData(rowNum, ws.Range(pmsColBaseMat & "1").Column)
                Exit For
            End If
        Next rowNum
        If i Mod 10 = 0 Then Application.StatusBar = "진행 중: " & i & " / " & tagList.Count & " 완료..."
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

    filePath = Application.GetOpenFilename("엑셀 파일(*.xls;*.xlsx;*.xlsb;*.xlsm), *.xls;*.xlsx;*.xlsb;*.xlsm", , "파일 선택", , False)

    If filePath = "False" Then
        MsgBox "취소", vbExclamation
        Set GetWorkbookFromUser = Nothing
        Exit Function
    End If

    fileName = Dir(filePath)

    isWorkbookOpen = False
    For Each wb In Workbooks
        If wb.Name = fileName Then
            isWorkbookOpen = True
            Set GetWorkbookFromUser = wb
            Exit Function
        End If
    Next wb

    If Not isWorkbookOpen Then
        Set wb = Workbooks.Open(filePath)
        Set GetWorkbookFromUser = wb
    End If

End Function

Function FindColLetter(hdr_row As Integer, search_value As Variant, Optional ws As Worksheet = Nothing) As String

    Dim search_rng As Range
    Dim found_cell As Range
    Dim col_letter As String

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
