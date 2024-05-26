Sub MergeExcelFiles()
    Dim FolderPath As String
    Dim FileName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wbDest As Workbook
    Dim wsDest As Worksheet
    Dim LastRow As Long
    Dim fd As FileDialog

    ' 폴더 선택을 위한 FileDialog 설정
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "폴더를 선택하세요"
        If .Show = -1 Then
            FolderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "폴더를 선택하지 않았습니다. 작업을 종료합니다."
            Exit Sub
        End If
    End With

    ' 새로운 워크북 생성
    Set wbDest = Workbooks.Add
    Application.DisplayAlerts = False
    While wbDest.Sheets.Count > 1
        wbDest.Sheets(1).Delete
    Wend
    Application.DisplayAlerts = True
    Set wsDest = wbDest.Sheets(1)
    wsDest.Name = "MergedData"

    ' 폴더 내 모든 엑셀 파일 불러오기
    FileName = Dir(FolderPath & "*.xls*")
    Do While FileName <> ""
        ' 소스 워크북 열기
        Set wbSource = Workbooks.Open(FolderPath & FileName)
        
        ' 각 시트 복사
        For Each wsSource In wbSource.Sheets
            wsSource.Copy After:=wbDest.Sheets(wbDest.Sheets.Count)
        Next wsSource
        
        ' 소스 워크북 닫기
        wbSource.Close SaveChanges:=False
        FileName = Dir
    Loop

    ' 기본 시트 삭제
    Application.DisplayAlerts = False
    wbDest.Sheets("MergedData").Delete
    Application.DisplayAlerts = True

    ' 결과 저장
    wbDest.SaveAs FolderPath & "MergedWorkbook.xlsx"
    wbDest.Close SaveChanges:=False

    MsgBox "모든 파일이 성공적으로 병합되었습니다!"
End Sub
