Sub MergebyPriority()

    
    Dim myPath As String
    Dim myFile As String
    Dim masterFilePath As String
    Dim subWorkbook As Workbook
    Dim mergeWorkbook As Workbook
    Dim fd As FileDialog
    Dim fileCount As Integer
    Dim processedCount As Integer

    Application.ScreenUpdating = False ' 화면 업데이트 끄기
    Application.DisplayAlerts = False ' 경고 표시 안함

    ' 폴더 선택 대화상자 초기화
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "폴더 선택"
        If .Show = -1 Then
            myPath = .SelectedItems(1) ' 폴더 선택
        Else
            MsgBox "선택 취소"
            Exit Sub
        End If
    End With

    ' 파일 선택 대화상자 초기화
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "취합 파일 선택"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls"
        If .Show = -1 Then
            masterFilePath = .SelectedItems(1) ' 마스터 파일 선택
        Else
            MsgBox "선택 취소"
            Exit Sub
        End If
    End With

    ' 마스터 워크북 열기
    Set mergeWorkbook = Workbooks.Open(masterFilePath)

    ' 파일 경로 검증
    If Right(myPath, 1) <> "\" Then myPath = myPath & "\"

    ' 파일 카운트 초기화
    myFile = Dir(myPath & "*.xls*")
    While myFile <> ""
        fileCount = fileCount + 1
        myFile = Dir
    Wend

    myFile = Dir(myPath & "*.xls*")
    Do While myFile <> ""
        ' 상태바 표시 업그레이드
        processedCount = processedCount + 1
        Application.StatusBar = "Processing file " & processedCount & " of " & fileCount ' 진행 상태 업데이트
        
        ' 개별 데이터시트 불러오기
        Set subWorkbook = Workbooks.Open(Filename:=myPath & myFile, ReadOnly:=True) ' 읽기 전용으로 열기
        
        MergeTables subWorkbook, mergeWorkbook
        
        On Error Resume Next
        subWorkbook.Close False ' 변경 사항 없으므로 저장하지 않고 닫음
        myFile = Dir ' 다음 파일로 이동
        On Error GoTo 0
    Loop
    
    mergeWorkbook.Save
    'mergeWorkbook.Close
    MsgBox "완료"

    Application.StatusBar = False ' 상태 바 초기화
    Application.ScreenUpdating = True ' 화면 업데이트
    Application.DisplayAlerts = True

End Sub
