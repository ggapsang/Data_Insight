Sub MakeMasterSheet()
    
    Dim fd As FileDialog
    Dim masterFilePath As String
    Dim masterWb As Workbook
    Dim masterWs As Worksheet
    Dim dataSheetFilePath As String
    Dim dataSheetWb As Workbook
    Dim names() As String
    Dim sheetNames() As String
    Dim i As Integer
    
    ' 마스터 파일 선택 대화상자 초기화
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "마스터 파일 선택"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls"
        If .Show = -1 Then
            masterFilePath = .SelectedItems(1) ' 마스터 파일 선택
        Else
            MsgBox "선택 취소"
            Exit Sub
        End If
    End With
    
    
    ' 마스터 파일 열기
    Set masterWb = Workbooks.Open(masterFilePath)
    Set masterWs = masterWb.Sheets(1)

    
    ' 데이터 시트 폼 선택 대화상자 초기화
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "데이터 시트 폼 선택"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls"
        If .Show = -1 Then
            dataSheetFilePath = .SelectedItems(1) '데이터시트 폼 선택
        Else
            MsgBox "선택 취소"
            Exit Sub
        End If
    End With
    
    ' 데이터 시트 파일 열기
    Set dataSheetWb = Workbooks.Open(dataSheetFilePath)
    
    ' 데이터 시트 파일의 이름 저장
    names = ExportNames(dataSheetWb)
    
    
    ' 저장된 이름을 마스터 시트에 입력
    
        For i = 1 To UBound(names)
        
            masterWs.Cells(1, i + 3).Value = names(i)
        
        Next i
        
    
    MsgBox "완료"
    dataSheetWb.Close SaveChanges:=False
    masterWb.Save
    
End Sub

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
        ExtractSheetName = "Sheet name not found"
    End If
    
End Function

