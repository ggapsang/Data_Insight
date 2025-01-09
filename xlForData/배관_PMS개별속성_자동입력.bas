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
    Dim pmsColCode As String, pmsColMin As String, pmsColMax As String
    Dim pmsColPwht As String, pmsColBaseMat As String

    pmsColCode = FindColLetter(1, "code", ws)
    pmsColMin = FindColLetter(1, "MIN (float)", ws)
    pmsColMax = FindColLetter(1, "MAX (float)", ws)
    pmsColPwht = FindColLetter(1, "GENERAL PWHT", ws)
    pmsColBaseMat = FindColLetter(1, "GENERAL BASE MATERIAL", ws)

    ' PMS 데이터를 Dictionary에 저장
    Dim pmsDict As Object
    Set pmsDict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        Dim key As String
        key = ws.Cells(i, ws.Range(pmsColCode & "1").Column).Value
        If Not pmsDict.exists(key) Then
            pmsDict.Add key, Array(ws.Cells(i, ws.Range(pmsColPwht & "1").Column).Value, ws.Cells(i, ws.Range(pmsColBaseMat & "1").Column).Value)
        End If
    Next i

    ' 개별 속성 데이터 저장
    Dim pipColTagNo As String, pipColGroupCode As String, pipColLineClass As String, pipColLineSize As String
    pipColTagNo = FindColLetter(1, "설비번호", pipWs)
    pipColGroupCode = FindColLetter(1, "속성 그룹 코드", pipWs)
    pipColLineClass = FindColLetter(1, "개별속성8", pipWs)
    pipColLineSize = FindColLetter(1, "개별속성9", pipWs)

    Dim visibleCells As Range, tagList As Collection
    Set tagList = New Collection

    Set visibleCells = pipWs.Range(pipWs.Cells(2, pipWs.Range(pipColGroupCode & "1").Column), pipWs.Cells(pipWs.Cells(pipWs.Rows.Count, pipWs.Range(pipColGroupCode & "1").Column).End(xlUp).Row, pipWs.Range(pipColGroupCode & "1").Column)).SpecialCells(xlCellTypeVisible)

    For Each visibleCell In visibleCells
        If InStr(1, visibleCell.Value, "03") > 0 Then
            tagList.Add Array(pipWs.Cells(visibleCell.Row, pipWs.Range(pipColLineClass & "1").Column).Value, pipWs.Cells(visibleCell.Row, pipWs.Range(pipColLineSize & "1").Column).Value)
        End If
    Next visibleCell

    ' 속성 데이터 입력
    Dim tag As Variant
    Application.StatusBar = "작업을 시작합니다..."
    For i = 1 To tagList.Count
        tag = tagList(i)
        If pmsDict.exists(tag(0)) Then
            Dim values As Variant
            values = pmsDict(tag(0))
            pipWs.Cells(i + 1, pipWs.Range(pipColPwht & "1").Column).Value = values(0)
            pipWs.Cells(i + 1, pipWs.Range(pipColBaseMat & "1").Column).Value = values(1)
        End If
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
