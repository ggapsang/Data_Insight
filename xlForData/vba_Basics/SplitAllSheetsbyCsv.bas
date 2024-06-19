Sub SaveAllSheetsAsCSV()
    Dim ws As Worksheet
    Dim csvFilePath As String
    Dim wb As Workbook
    Dim tempWb As Workbook
    
    ' 현재 워크북을 기준으로 CSV 파일 경로 설정
    csvFilePath = ThisWorkbook.Path & "\"
    
    ' 모든 시트를 반복
    For Each ws In ThisWorkbook.Sheets
        ' 임시 워크북 생성
        ws.Copy
        Set tempWb = ActiveWorkbook
        
        ' CSV 파일로 저장
        tempWb.SaveAs Filename:=csvFilePath & ws.Name & ".csv", FileFormat:=xlCSV, CreateBackup:=False
        
        ' 임시 워크북 닫기
        tempWb.Close SaveChanges:=False
    Next ws
    
    MsgBox "All sheets have been saved as CSV files."
End Sub
