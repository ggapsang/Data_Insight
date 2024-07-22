Sub GetInstrumentAttribute()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Starting the process..."
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim mappingWs As Worksheet
    
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("계기")
    Set mappingWs = wb.Sheets("표준데이터시트 매핑")
    
    Dim wsLastRow As Long
    Dim dirName As String
    
    Dim colDirName As String
    Dim colExtractComplet As String
    Dim colFormNm As String
    Dim colAttrGroupCode As String

    wsLastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    mappingWsLastRow = mappingWs.Cells(Rows.Count, 1).End(xlUp).Row
    colDirName = FindColLetter(1, "Directory", ws)
    colExtractComplet = FindColLetter(1, "추출 완료", ws)
    colFormNm = FindColLetter(1, "타입(폼명)", ws)
    colAttrGroupCode = FindColLetter(1, "속성 그룹 코드", ws)

    Dim rngMappingWs As Range
    Set rngMappingWs = mappingWs.Range("A1:D" & mappingWsLastRow)
    
''' 속성 그룹코드가 03_DATA이고, K열이 비어 있을 때, rngDir의 값을 순회하면서 파일을 하나씩 열고, 시트 매핑에 따라 추출 시작

    Dim i As Long

    For i = 2 To wsLastRow
        Application.StatusBar = "Processing row " & i & " of " & wsLastRow

        If IsEmpty(ws.Range(colExtractComplet & i).Value) And ws.Range(colAttrGroupCode & i) = "03_DATA" Then
            Dim file_path As String
            Dim datasheetWb As Workbook

            file_path = ws.Range(colDirName & i).Value
            Set datasheetWb = Workbooks.Open(file_path)

            Dim typeNm As String
            typeNm = ws.Range(colFormNm & i).Value
            
            ' 표준데이터시트 매핑 시트에서 타입으로 필터링 후 E 열에 해당하는 위치 값에, D열 값 넣기
            rngMappingWs.AutoFilter Field:=1, Criteria1:=typeNm

            ' 필터링으로 보이는 값에 대해서만 순회

            Dim cell As Range
            Dim rowNumber As Long
            Dim namedRange As String
            Dim subnamedRange As String
            Dim extractValue As String

            On Error Resume Next
            For Each cell In rngMappingWs.Columns(1).SpecialCells(xlCellTypeVisible).Cells
                If cell.Row > 1 Then '해더 행 제외
                    rowNumber = cell.Row
                    colLetter = mappingWs.Range("E" & rowNumber).Value
                    namedRange = mappingWs.Range("D" & rowNumber).Value
                    subnamedRange = mappingWs.Ragne("F" & rowNumber).Value

                    extractValue = GetNamedRangeValue(namedRange, datasheetWb)
                    If extractValue = "Error: not defined and invalid format" Then
                        extractValue = GetNamedRangeValue(subnamedRange, datasheetWb)
                    End If

                    If InStr(1, visibleCell.Value, "NOTE") > 0 Or InStr(1, visibleCell.Value, "Note") Then
                        extractValue = GetNamedRangeValue(subnamedRange, datasheetWb)
                    End If

                    ws.Range(colLetter & i).Value = extractValue
    
                End If
            Next cell
            On Error GoTo 0

            ' 오토필터 해제, 데이터시트 종료
            mappingWs.AutoFilterMode = False
            datasheetWb.Close SaveChanges:=False
        End If
    
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.StatusBar = False ' 상태 표시줄 초기화


End Sub


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
    End If

End Function

Function GetNamedRangeValue(namedRange As String, wb As Workbook) As Variant

    Dim rng As Range
    Dim sheetName As String
    Dim cellAddress As String
    Dim exclamationPos As Integer

    On Error GoTo ErrHandler
    ' "N/A" 값이 인자로 들어왔을 경우 아무것도 리턴하지 않음
    If namedRange = "N/A" Then
        GetNamedRangeValue = ""
        Exit Function
    End If

    '정의된 이름이 존재하는지 확인
    On Error Resume Next
    Set rng = wb.Names(namedRange).RefersToRange
    On Error GoTo 0

    If Not rng Is Nothing Then
        '정의된 이름에 해당하는 범위를 가져옴
        GetNamedRangeValue = rng.Cells(1, 1).Value
    Else
        ' 시트명과 셀 주소 형식인지 확인
        exclamationPos = InStr(namedRange, "!")
        If exclamationPos > 0 Then
            sheetName = Left(namedRange, exclamationPos - 1)
            cellAddress = Mid(namedRange, exclamationPos + 1)
            
            On Error Resume Next
            Set rng = wb.Sheets(sheetName).Range(cellAddress)
            On Error GoTo 0
            
            If Not rng Is Nothing Then
                ' 시트명과 셀 주소에 해당하는 범위를 가져옴
                GetNamedRangeValue = rng.Cells(1, 1).Value
            Else
                GetNamedRangeValue = "Error: Invalid sheet name or cell address"
            End If
        Else
            GetNamedRangeValue = "Error: not defined and invalid format"
        End If
    End If

    Exit Function

ErrHandler:
    ' 오류가 발생하면 에러 메시지를 리턴
    GetNamedRangeValue = "Error: " & Err.Description

End Function


