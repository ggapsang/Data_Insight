Sub 배관개별속성입력()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

'' 전역 변수

'''워크시트
    
    'PMS 시트 변수 정의
    
        Dim wb As Workbook
        Dim ws As Worksheet
    
        Set wb = ThisWorkbook
        Set ws = wb.Sheets("PMS")

        '사이즈 매핑 시트 정의
        Dim sizeWs As Worksheet
        Set sizeWs = wb.Sheets("사이즈 변환")
    

    ' 개별속성을 입력할 워크시트
        Dim pipWb As Workbook
        Dim pipWs As Worksheet
        Dim pipWsNm As String
    
        Set pipWb = GetWorkbookFromUser

        pipWsNm = InputBox("워크시트 이름")

    On Error Resume Next ' 오류가 발생하면 다음 줄로 넘어감
        Set pipWs = pipWb.Sheets(pipWsNm)
        If Err.Number <> 0 Then ' 워크시트를 찾을 수 없으면 오류 번호가 설정됨
            MsgBox "워크시트 '" & pipWsNm & "'을(를) 찾을 수 없습니다.", vbExclamation
            Set pipWs = Nothing
        End If
    On Error GoTo 0 ' 기본 오류 처리로 돌아감
    
    
'' PMS 칼럼, 마지막 행

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
    pmsColMat = FindColLetter(1, "GENERAL MATERIAL", ws)
    pmsColRating = FindColLetter(1, "GENERAL RATING", ws)
    pmsColConnectType = FindColLetter(1, "GENERAL END CONNECTION TYPE", ws)
    pmsColSch = FindColLetter(1, "GENERAL SCHEDULE", ws)
    pmsColNdt = FindColLetter(1, "GENERAL NON DESTRUCTIVE TEST RATE", ws)
    
    
    Dim pmsLastRow As Integer
    
    pmsLastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

'' 개별속성 리스트 시트에 정보들을 배열에 저장

    '' 메인 해더 칼럼 변수 확인
    
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
        pipColLineClass = FindColLetter(1, "개별속성8", pipWs) '개별속성8 : GENERAL LINE CLASS
        pipColLineSize = FindColLetter(1, "개별속성9", pipWs) ' 개별속성9 : GENERAL LINE SIZE
        pipColPwht = FindColLetter(1, "개별속성19", pipWs) '개별속성19 : GENERAL PWHT
        pipColBaseMat = FindColLetter(1, "개별속성20", pipWs) '개별속성20 : GENERAL BASE MATERIAL
        pipColCA = FindColLetter(1, "개별속성21", pipWs) '개별속성21 : CORROSION ALLOWANCE
        pipColRating = FindColLetter(1, "개별속성22", pipWs) '개별속성22 : GENERAL RATING
        pipColMaterial = FindColLetter(1, "개별속성23", pipWs) '개별속성23 : GENERAL MATERIAL
        pipColConType = FindColLetter(1, "개별속성25", pipWs) '개별속성25 : GENERAL END CONNECTION TYPE
        pipColSch = FindColLetter(1, "개별속성26", pipWs) '개별속성26 : GENERAL SCHEDULE
        pipColNDE = FindColLetter(1, "개별속성27", pipWs) '개별속성27 : GENERAL NON DESTRUCTIVE TEST RATE

    '' 정렬
        Dim pipLastRow As Long
        Dim pipLastCol As Long

        pipLastRow = pipWs.Cells(1, 1).End(xlDown).Row
        pipLastCol = pipWs.Cells(1, 1).End(xlToRight).Column



    'visible range 중에서 속성 그룹 코드에 '03'이라는 문자열을 포함한 값들의 tagNo(pipColTagNo가 있는 곳)만 배열로 저장
    ' 이 배열은 태그 번호를 1차원으로, 그 다음에 개별속성1, 개별속성2를 그 하위로 가지는 배열로 생성됨

        Dim visibleCell As Range
        Dim TagList As Collection
        Set TagList = New Collection

        For Each visibleCell In pipWs.Range(pipWs.Cells(2, pipWs.Range(pipColGroupCode & "1").Column), pipWs.Cells(pipLastRow, pipWs.Range(pipColGroupCode & "1").Column)).SpecialCells(xlCellTypeVisible)
            If InStr(1, visibleCell.Value, "03") > 0 Then ' '03'을 포함하는지 확인
                ' 설비번호와 관련 속성들을 배열에 저장
                Dim tagData(1 To 2) As Variant
                Dim strSize As String
                strSize = pipWs.Cells(visibleCell.Row, pipWs.Range(pipColLineSize & "1").Column).Value
                tagData(1) = pipWs.Cells(visibleCell.Row, pipWs.Range(pipColLineClass & "1").Column).Value
                tagData(2) = GetFloatSize(strSize, sizeWs)
                TagList.Add tagData
            End If
        Next visibleCell

    ' Collection을 배열로 변환
        Dim Tags() As Variant
        ReDim Tags(1 To TagList.Count, 1 To 2)
        Dim i As Integer
        i = 1

        Dim item As Variant
        For Each item In TagList
            Tags(i, 1) = item(1)
            Tags(i, 2) = item(2)
            i = i + 1
        Next item


''PMS 시트의 정보들을 배열에 저장

    Dim startRow As Long
    
    startRow = InputBox("시작 행 번호", "시작 행 입력")
    
    
    startRow = 6
    Dim PMS() As Variant
    ReDim PMS(1 To pmsLastRow, 1 To 11)
    
    Dim l As Long

    For l = 2 To pmsLastRow - startRow + 1
        PMS(l - 1, 1) = ws.Cells(l, ws.Range(pmsColCode & "1").Column).Value 'PMS CODE
        PMS(l - 1, 2) = ws.Cells(l, ws.Range(pmsColMin & "1").Column).Value 'MIN (float)
        PMS(l - 1, 3) = ws.Cells(l, ws.Range(pmsColMax & "1").Column).Value 'MAS (float)
        PMS(l - 1, 4) = ws.Cells(l, ws.Range(pmsColPwht & "1").Column).Value 'PWHT
        PMS(l - 1, 5) = ws.Cells(l, ws.Range(pmsColBaseMat & "1").Column).Value 'GENERAL BASE MATERIAL
        PMS(l - 1, 6) = ws.Cells(l, ws.Range(pmsColCa & "1").Column).Value 'GENERAL C.A(mm)
        PMS(l - 1, 7) = ws.Cells(l, ws.Range(pmsColRating & "1").Column).Value 'GENERAL RATING
        PMS(l - 1, 8) = ws.Cells(l, ws.Range(pmsColMat & "1").Column).Value 'GENERAL MATERIAL
        PMS(l - 1, 9) = ws.Cells(l, ws.Range(pmsColConnectType & "1").Column).Value 'GENERAL END CONNECTION TYPE
        PMS(l - 1, 10) = ws.Cells(l, ws.Range(pmsColSch & "1").Column).Value 'GENERAL SCHEDULE
        PMS(l - 1, 11) = ws.Cells(l, ws.Range(pmsColNdt & "1").Column).Value 'GENERAL NON DESTRUCTIVE TEST RATE
    Next l


'' Tag의 속성 항목들을 순회하면서 데이터 입력. 첫 번째 키는 일치하고, 두 번째 키는 범위 내에 있어야 함
    Dim k As Long
    Dim idx As Integer

    Application.StatusBar = "작업을 시작합니다..."
    For k = 1 To pipLastRow - startRow + 1
    
        Dim strKey As String
        Dim numberKey As Double
        Dim searchResults(4 To 11) As Variant
    
        strKey = Tags(k, 1)
        numberKey = Tags(k, 2)
    
        ' SearchValue 호출을 한 번으로 줄이고 결과 캐싱
        For i = 4 To 11
            searchResults(i) = SearchValue(strKey, numberKey, i, PMS)
        Next i
    
        ' With 구문을 사용하여 pipWs 호출 최소화
        With pipWs
            .Range(pipColPwht & startRow + k - 1).Value = searchResults(4) ' PWHT 입력
            .Range(pipColBaseMat & startRow + k - 1).Value = searchResults(5) ' BASE MATERIAL 입력
            .Range(pipColCA & startRow + k - 1).Value = searchResults(6) ' C.A 입력
            .Range(pipColRating & startRow + k - 1).Value = searchResults(7) ' RATING 입력
            .Range(pipColMaterial & startRow + k - 1).Value = searchResults(8) ' MATERIAL 입력
            .Range(pipColConType & startRow + k - 1).Value = searchResults(9) ' END CONNECTION TYPE 입력
            .Range(pipColSch & startRow + k - 1).Value = searchResults(10) ' SCHEDULE 입력
            .Range(pipColNDE & startRow + k - 1).Value = searchResults(11) ' NDE RATE 입력
        End With
    
        ' 상태 표시줄 업데이트
        If k Mod 10 = 0 Then ' 10개의 행마다 상태 업데이트
            Application.StatusBar = "처리 중: " & k & " / " & Total & " (" & Format(k / pipLastRow, "0%") & " 완료)"
        End If
    
    Next k
    
    MsgBox "완료"
   
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.StatusBar = False

End Sub



Function SearchValue(strKey As String, numberKey As Double, idxNo As Integer, arrValue As Variant) As String
    
    Dim i As Long
    
    For i = LBound(arrValue, 1) To UBound(arrValue, 1)
        If strKey = arrValue(i, 1) Then 'PMS CODE 가 일치할 때
            If numberKey >= arrValue(i, 2) And numberKey <= arrValue(i, 3) Then
                SearchValue = arrValue(i, idxNo)
                Exit Function
            End If
        End If
    Next i
    SearchValue = ""

End Function



Function GetFloatSize(strSize As String, Optional ws As Worksheet) As Double

    If ws Is Nothing Then
        Set ws = ActiveWorkbook.ActiveSheet
    End If


    Dim lastRow As Integer
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim keyData As Variant
    Dim valueData As Variant
    keyData = ws.Range("A2:A" & lastRow).Value
    valueData = ws.Range("B2:B" & lastRow).Value

    Dim strSizeList As Variant
    Dim floatSizeList As Variant

    strSizeList = keyData
    floatSizeList = valueData

    Dim i As Integer
    For i = LBound(strSizeList, 1) To UBound(strSizeList, 1)
        If strSize = strSizeList(i, 1) Then
            GetFloatSize = floatSizeList(i, 1)
            Exit Function
        End If
    Next i

    GetFloatSize = 0

End Function



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
        MsgBox (search_value & "is not found")
    End If

End Function
