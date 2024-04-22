''' Function_xlNames.bas의 함수들 참조 필요
Sub Main()
    
    Dim myPath As String
    Dim myFile As String
    Dim masterFilePath As String
    Dim myWorkbook As Workbook
    Dim masterWorkbook As Workbook
    Dim masterSheet As Worksheet
    Dim fd As FileDialog
    Dim fileCount As Integer
    Dim processedCount As Integer
    Dim charstartRow As String
    Dim startRow As Long
       
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
    
    
    Application.ScreenUpdating = False ' 화면 업데이트 끄기
    Application.DisplayAlerts = False ' 경고 표시 안함
    
    
    ' 마스터 파일 열기
    Set masterWorkbook = Workbooks.Open(masterFilePath)
    Set masterSheet = masterWorkbook.Sheets(1)


    '몇 번째 행부터 입력을 다시 시작할까요
    charstartRow = InputBox("몇 번째 행부터 입력을 다시 시작할까요")
    startRow = CLng(charstartRow)
    
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
        Set myWorkbook = Workbooks.Open(Filename:=myPath & myFile, ReadOnly:=True) ' 읽기 전용으로 열기
        
        FindValuesAndMove_Rev myWorkbook, masterSheet, startRow
        
        On Error Resume Next
        myWorkbook.Close False ' 변경 사항 없으므로 저장하지 않고 닫음
        myFile = Dir ' 다음 파일로 이동
        On Error GoTo 0
    Loop

    Application.StatusBar = False ' 상태 바 초기화
    Application.ScreenUpdating = True ' 화면 업데이트
    Application.DisplayAlerts = True


    MsgBox "완료"
    masterWorkbook.Save
        
End Sub


Public Sub FindValuesAndMove_Rev(sourceWorkbook As Workbook, masterSheet As Worksheet, startRow As Long)
    
    Dim sourceSheet As Worksheet
    Dim cellAddress As String
    Dim sourceAddress() As String
    Dim i As Long, j As Long
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim fileRow As Long
    Dim beginColumn As Long
    Dim workbookName As String
    Dim wokrsheetName As String
    Dim check_flag As Boolean


    ' 마스터 파일에서 데이터를 찾을 마지막 행 결정
    lastRow = masterSheet.Cells(masterSheet.Rows.Count, "B").End(xlUp).Row
    
    ' 마스터 파일의 마지막 열 결정
    lastColumn = masterSheet.Cells(1, Columns.Count).End(xlToLeft).Column

    ' 소스 파일의 이름 결정
    workbookName = sourceWorkbook.Name
    
    ' 마스터 파일의 두번째 행부터 마지막 행까지 순회하면서 소스 파일(데이터시트)이 있는 워크북의 행 번호 찾기
    fileRow = 0
    For i = startRow To lastRow
        Dim fileNameCell As String
        fileNameCell = masterSheet.Range("B" & i).Value
        
        
        If fileNameCell = workbookName Then
            fileRow = i
            Exit For ' 파일을 찾으면 반복 중지
        End If
    Next i

    If fileRow = 0 Then
        'Debug.Print sourceWorkbook.Name
        Exit Sub ' 파일을 찾지 못했으므로 서브 프로시저 종료
    End If


    ' 데이터가 들어가는 열(beginColumn)부터 시작하여 마지막 열까지 순회
    beginColumn = 4
        
    For j = beginColumn To lastColumn
        cellAddress = masterSheet.Cells(1, j).Value
        sourceAddress = Split(cellAddress, "!")
        worksheetName = sourceWorkbook.Sheets(1).Name
            
        If UBound(sourceAddress) > 0 Then
        
            If sourceAddress(0) = worksheetName Then
                
                Set sourceSheet = sourceWorkbook.Worksheets(sourceAddress(0))
            
            Else
                Set sourceSheet = sourceWorkbook.Worksheets(worksheetName)
            
            End If
        
        End If
          
        On Error Resume Next
        masterSheet.Cells(fileRow, j).Value = sourceSheet.Range(sourceAddress(1)).Value
        On Error GoTo 0
        Debug.Print "Error" & sourceAddress(1)
    
    Next j
                     
    sourceWorkbook.Close SaveChanges:=False
    
End Sub
