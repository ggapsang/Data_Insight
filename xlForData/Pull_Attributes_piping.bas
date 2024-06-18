Sub Okinawa()

''기본옵션
    
    '작동시간 측정
    Dim StartTime As Date
    StartTime = Timer

    '매크로 작업 중 스크린 활성화 정지
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    '작업 파일 관련 변수 선언
    Dim tgtWb As Workbook
    Dim tgtWs As Worksheet
    Dim tgtFile As Variant
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    
''파일 불러오기(소스파일, 타겟파일)
    'source worksheet(PMS 테이블) 세팅
    Set srcWb = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*", Title:="불러올 파일 선택", MultiSelect:=False)
    If tgtFile = False Then
        MsgBox "취소"
        Exit Sub
    End If
    Set srcWs = srcWb.Worksheets("source_data")
    
    ' target workbook 세팅 명령창 실행
    tgtFile = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*", Title:="불러올 파일 선택", MultiSelect:=False)
    If tgtFile = False Then
        MsgBox "취소"
        Exit Sub
    End If
    
    ' target workbook이 이미 열려있을 경우
    Set tgtWb = GetWorkbook(tgtFile)

    ' target workbook이 열려 있지 않을 경우 해당 파일을 열음
    If tgtWb Is Nothing Then
        Set tgtWb = Workbooks.Open(tgtFile)
    End If

    ' target worksheet 세팅
    Set tgtWs = tgtWb.ActiveSheet
    tgtWs.Activate
    
    ' target worksheet의 주요 칼럼들의 열 번호(알파벳) 변수 선언
    Dim tgtMatCol As String
    Dim tgtFluidCol As String
    Dim tgtSerialCol As String
    Dim tgtSizeCol As String
    Dim tgtInsulCol As String
    Dim tgtTracingCol As String
    
    Dim tgtDwgNoCol As String
    Dim tgtisoNoCol As String
    
    Dim tgtTypeCol As String
    Dim tgtDescrCol As String
    
    Dim tgtFluidValueCol As String
    Dim tgtSerialValueCol As String
    Dim tgtMatValueCol As String
    
    Dim tgtDwgNoValueCol As String
    Dim tgtIsoNoValueCol As String
    
    Dim tgtFlngRatingCol As String
    Dim tgtFlngFaceCol As String
    Dim tgtPipeMatCol As String
    Dim tgtGskCol As String
    Dim tgtNpsCol As String
    Dim tgtOpPresCol As String
    Dim tgtOpTempCol As String
    Dim tgtDesignPresCol As String
    Dim tgtDesignTempCol As String
    Dim tgtTcTypeCol As String
    Dim tgtFluidPhaseCol As String
    Dim tgtNdeCol As String
    Dim tgtPwhtCol As String
    Dim tgtPidNoCol As String
    Dim tgtIsoCol As String
    Dim tgtSchCol As String
    Dim tgtWtCol As String
    Dim tgtInsTypeCol As String
    
    ' 해더 행 번호 설정
    Dim hdrRow As Integer
    hdrRow = 1
    
    ' 마지막 행 번호 찾기(A열을 기준으로 카운팅)
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).row
    
    
    ' target worksheet의 주요 칼럼들의 열 번호(알파벳) 찾기

    
    tgtMatCol = FindColLetter(hdrRow, "배관 사양 코드")
    tgtFluidCol = FindColLetter(hdrRow, "사용 유체 코드")
    tgtSerialCol = FindColLetter(hdrRow, "태그 시리얼 번호")
    tgtSizeCol = FindColLetter(hdrRow, "배관 사이즈")
    tgtInsulCol = FindColLetter(hdrRow, "배관 보온 코드")
    tgtTracingCol = FindColLetter(hdrRow, "배관 트레이싱 코드")
    
    tgtDwgNoCol = FindColLetter(hdrRow, "도면번호")
    tgtisoNoCol = FindColLetter(hdrRow, "ISO 도면번호")
    
    tgtTypeCol = FindColLetter(hdrRow, "타입 이름")
    tgtDescrCol = FindColLetter(hdrRow, "MDM 설비 내역")
    
    tgtDwgNoValueCol = FindColLetter(hdrRow, "P&ID 번호")
    tgtIsoNoValueCol = FindColLetter(hdrRow, "ISO DWG 번호")
    
    tgtFluidValueCol = FindColLetter(hdrRow, "사용 유체")
    tgtSerialValueCol = FindColLetter(hdrRow, "시리얼 번호")
    tgtMatValueCol = FindColLetter(hdrRow, "재질 클래스")
    tgtFlngRatingCol = FindColLetter(hdrRow, "플랜지 래이팅")
    tgtFlngFaceCol = FindColLetter(hdrRow, "플랜지 접촉면 타입")
    tgtPipeMatCol = FindColLetter(hdrRow, "배관 재질")
    tgtGskCol = FindColLetter(hdrRow, "가스켓 재질")
    tgtNpsCol = FindColLetter(hdrRow, "표준 배관 구경")
    tgtOpPresCol = FindColLetter(hdrRow, "운전압력")
    tgtOpTempCol = FindColLetter(hdrRow, "운전온도")
    tgtDesignPresCol = FindColLetter(hdrRow, "설계압력")
    tgtDesignTempCol = FindColLetter(hdrRow, "설계온도")
    tgtTcTypeCol = FindColLetter(hdrRow, "트레이싱 타입")
    tgtFluidPhaseCol = FindColLetter(hdrRow, "유체 상태")
    tgtNdeCol = FindColLetter(hdrRow, "비파괴 검사율")
    tgtPwhtCol = FindColLetter(hdrRow, "후열처리 여부")
    tgtPidNoCol = FindColLetter(hdrRow, "P&ID 번호")
    tgtIsoCol = FindColLetter(hdrRow, "ISO DWG 번호")
    tgtSchCol = FindColLetter(hdrRow, "배관 스케줄")
    tgtWtCol = FindColLetter(hdrRow, "배관 두께")
    tgtInsTypeCol = FindColLetter(hdrRow, "보온 타입")
    
    
    ''' 사용 유체, 시리얼 번호, 재질 클래스, 표준 배관 구경, 트레이싱 타입, 보온 타입, P&ID 번호, ISO DWG 번호 값 넣기
    Dim row As Long
    For row = hdrRow + 1 To lastRow
        Range(tgtFluidValueCol & row).Value = Range(tgtFluidCol & row).Value '사용 유체
        Range(tgtSerialValueCol & row).Value = Range(tgtSerialCol & row).Value '시리얼 번호
        Range(tgtMatValueCol & row).Value = Range(tgtMatCol & row).Value '재질 클래스
        Range(tgtNpsCol & row).Value = Left(Range(tgtSizeCol & row).Value, Len(Range(tgtSizeCol & row).Value) - 1) & "|in" '표준 배관 구경
        Range(tgtTcTypeCol & row).Value = Range(tgtTracingCol & row).Value '트레이싱 타입
        Range(tgtInsTypeCol & row).Value = Range(tgtInsulCol & row).Value '배관 보온 타입
        Range(tgtDwgNoValueCol & row).Value = Range(tgtDwgNoCol & row).Value 'P&ID 번호
        Range(tgtIsoNoValueCol & row).Value = Range(tgtisoNoCol & row).Value 'ISO DWG 번호
    Next row
    
    ''' 타입 입력
    srcWs.Activate
    
    For row = hdrRow + 1 To lastRow
        On Error GoTo InvalidValue_cct:
        tgtWs.Range(tgtTypeCol & row) = Application.WorksheetFunction.VLookup(tgtWs.Range(tgtMatCol & row), Range("A:D"), 4, False)
    Next row
    
    
    ''' 플랜지 래이팅 입력
    For row = hdrRow + 1 To lastRow
        On Error GoTo InvalidValue_flngRating:
        tgtWs.Range(tgtFlngRatingCol & row) = Application.WorksheetFunction.VLookup(tgtWs.Range(tgtMatCol & row), Range("A:G"), 5, False)
    Next row
    
    
    ''' 플랜지 접촉면 타입 입력
    For row = hdrRow + 1 To lastRow
        On Error GoTo invalidValue_flngFace
    tgtWs.Range(tgtFlngFaceCol & row) = Application.WorksheetFunction.VLookup(tgtWs.Range(tgtMatCol & row), Range("A:F"), 6, False)
    Next row
    
    
    ''' 배관 재질, 가스켓 재질, 운전압력, 운전온도, 설계압력, 설계온도, 비파괴 검사율, 후열처리 여부, 배관 스케줄 값 넣기
    Dim sizeValue As Single
    Dim minSize As Single
    Dim maxSize As Single
    Dim signal As Boolean
    Dim lastRowPms As Integer
    
    lastRowPms = Cells(Rows.Count, 1).End(xlUp).row
    
    For row = hdrRow + 1 To lastRow
    On Error GoTo InvalidAction:
        
        sizeValue = CSng(Left(tgtWs.Range(tgtSizeCol & row).Value, Len(tgtWs.Range(tgtSizeCol & row).Value) - 1))
            
        Dim row_2 As Long
        For row_2 = 2 To lastRowPms
            minSize = CSng(Range("B" & row_2))
            maxSize = CSng(Range("C" & row_2))
            
            If tgtWs.Range(tgtMatCol & row).Value = Range("A" & row_2) And IsInRange(sizeValue, minSize, maxSize) = True Then
                tgtWs.Range(tgtPipeMatCol & row).Value = Range("G" & row_2).Value
                tgtWs.Range(tgtGskCol & row).Value = Range("H" & row_2).Value
                
                temp = checkColorAndFill(tgtWs.Range(tgtOpPresCol & row), Range("I" & row_2))
                temp = checkColorAndFill(tgtWs.Range(tgtOpTempCol & row), Range("J" & row_2))
                temp = checkColorAndFill(tgtWs.Range(tgtDesignPresCol & row), Range("K" & row_2))
                temp = checkColorAndFill(tgtWs.Range(tgtDesignTempCol & row), Range("L" & row_2))
                temp = checkColorAndFill(tgtWs.Range(tgtNdeCol & row), Range("M" & row_2))
                temp = checkColorAndFill(tgtWs.Range(tgtPwhtCol & row), Range("N" & row_2))
                    
                If Range("O" & row_2).Value = "CALC" Then
                    tgtWs.Range(tgtWtCol & row) = Range("P" & row_2).Value
                Else
                    tgtWs.Range(tgtSchCol & row) = Range("O" & row_2).Value
                End If
            Else
            End If
        Next row_2
    Next row
    
    
    ''' 배관 두께 입력
    Dim srcWs2 As Worksheet
    Set srcWs2 = srcWb.Worksheets("배관 스케줄")
    
    srcWs2.Activate
    
    Dim lastRowSch As Integer
    Dim joinKey As String
    
    lastRowSch = Cells(Rows.Count, 1).End(xlUp).row
    
    
    For row = hdrRow + 1 To lastRowSch
    
        If tgtWs.Range(tgtWtCol & row).Value = "" Then
            joinKey = tgtWs.Range(tgtSizeCol & row).Value & "-" & tgtWs.Range(tgtSchCol & row)
            For row_2 = 2 To lastRowSch
                If Range("B" & row_2).Value = joinKey Then
                    tgtWs.Range(tgtWtCol & row).Value = Range("F" & row_2).Value
                Else
                End If
            Next row_2
        Else
        End If
    Next row
    
   
    ''' MDM 설비 내역 입력
    Dim srcWs3 As Worksheet
    Set srcWs3 = srcWb.Worksheets("사용 유체 코드")
    
    srcWs3.Activate
    
    Dim lasRowFluid As Integer
    
    lastRowFluid = Cells(Rows.Count, 1).End(xlUp).row
    
    For row = hdrRow + 1 To lastRow
        tgtWs.Range(tgtDescrCol & row).Value = tgtWs.Range(tgtFluidCol & row).Value
        
        For row_2 = 2 To lastRowFluid
            If tgtWs.Range(tgtFluidCol & row).Value = Range("A" & row_2).Value Then
                tgtWs.Range(tgtDescrCol & row).Value = Range("C" & row_2).Value
            Else
            End If
        Next row_2
    Next row
    
    MsgBox prompt:=Format(Timer - StartTime, "0.000") & "sec"
    

Exit Sub

InvalidValue_cct:
    tgtWs.Range(tgtTypeCol & row) = "TBD"
    
    Resume Next
    
InvalidValue_flngRating:
    tgtWs.Range(tgtFlngRatingCol & row) = ""
    
    Resume Next

invalidValue_flngFace:
    tgtWs.Range(tgtFlngFaceCol & row) = ""

InvalidAction:

    Resume Next

End Sub

Function checkColorAndFill(fill As Range, src As Range) As Integer

    If src.Interior.ColorIndex = "15" Then
    checkColorAndFill = 0
    Else
    fill.Value = src.Value
    checkColorAndFill = 0
    End If

End Function

Function IsInRange(sizeValue As Single, minSize As Single, maxSize As Single) As Boolean

    If sizeValue >= minSize And sizeValue <= maxSize Then
        IsInRange = True
    Else
        IsInRange = False
    End If

End Function

Function GetWorkbook(ByVal sFullName As String) As Workbook
    Dim sFile As String
    Dim wbReturn As Workbook

    sFile = Dir(sFullName)

    On Error Resume Next
    Set wbReturn = Workbooks(sFile)

    If wbReturn Is Nothing Then
        Set GetWorkbook = Nothing
    Else
        Set GetWorkbook = wbReturn
    End If

    On Error GoTo 0
End Function

Function FindColLetter(hdr_row As Integer, search_value As Variant) As String

    Dim search_rng As Range
    Dim found_cell As Range
    Dim col_letter As String
    
    Set search_rng = ActiveSheet.Rows(hdr_row)
    
    Set found_cell = search_rng.Find(What:=search_value, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not found_cell Is Nothing Then
        col_letter = Replace(found_cell.Cells.Address(False, False), hdr_row & "", "")
        FindColLetter = col_letter
    Else
        FindColLetter = InputBox(search_value & "칼럼을 찾을 수 없습니다.해당 칼럼의 열 번호(알파벳,대소문자 구분필수)를 직접 입력하세요")

    End If

End Function
