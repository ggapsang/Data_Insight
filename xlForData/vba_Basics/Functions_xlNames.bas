''' 워크북에 정의된 이름들을 배열에 모두 저장
Function ExportNames(wb As Workbook)
    
    Dim i As Integer
    Dim namesArray() As String
    Dim nm As String
    Dim nm_ref As String
    Dim nm_sheetNm As String
    
    ' 이름을 저장할 배열 초기화
    ReDim namesArray(1 To wb.names.Count)
    
    ' 모든 이름을 배열에 저장
    For i = 1 To wb.names.Count
        
        nm = wb.names(i).Name
        nm_ref = wb.names(i).RefersTo
        nm_sheetNm = ExtractSheetName(nm_ref)
        
        namesArray(i) = nm_sheetNm & "!" & nm
    
    Next i
    
    ExportNames = namesArray

End Function

''' 이름에 참조된 모든 셀 주소를 배열에 저장
Function ExportNmRef(wb as Wokrbook)

    Dim i As Integer
    Dim namesArrayRef() As String
    Dim nm As String
    Dim nm_ref As String
    Dim nm_sheetNm As String
    
    ' 이름을 저장할 배열 초기화
    ReDim namesArray(1 To wb.names.Count)
    
    ' 모든 이름을 배열에 저장
    For i = 1 To wb.names.Count
        
        nm = wb.names(i).Name
        nm_ref = wb.names(i).RefersTo
        
        namesRefArray(i) = nm_ref
    
    Next i
    
    ExportNames = namesArrayRef

End Function

''' 참조 주소에서 시트 이름만 추출
Function ExtractSheetName(ref As String) As String
    Dim exclamPos As Integer
    Dim sheetName As String
    
    ' 참조 문자열에서 '!' 위치를 찾음
    exclamPos = InStr(ref, "!")
    
    ' '!'가 있는 경우, 그 앞의 문자열을 추출
    If exclamPos > 0 Then
        sheetName = Left(ref, exclamPos - 1)
        ' 작은따옴표 제거
        sheetName = Replace(sheetName, "'", "")
        sheetName = Replace(sheetName, "=", "")
        ExtractSheetName = sheetName
    Else
        '오류가 있을 경우 아래 메세지를 출력
        ExtractSheetName = "Sheet name not found"
    End If
    
End Function

''' 이름이 이미 존재하는지 확인
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
