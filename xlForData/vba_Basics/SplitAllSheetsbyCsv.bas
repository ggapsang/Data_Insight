Sub SaveAllSheetsAsCSV()
    Dim ws As Worksheet
    Dim csvFolderPath As String
    Dim wb As Workbook
    Dim tempWb As Workbook
    Dim folderName As String
    Dim fso As Object
    Dim fileName As String
    
    ' FileSystemObject 생성
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 현재 워크북의 파일 이름을 확장자 없이 가져옴
    fileName = ThisWorkbook.Name
    folderName = fso.GetBaseName(fileName)
    
    ' 폴더 경로 설정
    csvFolderPath = ThisWorkbook.Path & "\" & folderName & "\"
    
    ' 폴더가 존재하지 않으면 생성
    If Dir(csvFolderPath, vbDirectory) = "" Then
        MkDir csvFolderPath
    End If
    
    ' 모든 시트를 반복
    For Each ws In ThisWorkbook.Sheets
        ' 임시 워크북 생성
        ws.Copy
        Set tempWb = ActiveWorkbook
        
        ' CSV 파일로 저장
        tempWb.SaveAs Filename:=csvFolderPath & ws.Name & ".csv", FileFormat:=xlCSV, CreateBackup:=False
        
        ' 임시 워크북 닫기
        tempWb.Close SaveChanges:=False
    Next ws
    
    MsgBox "All sheets have been saved as CSV files in folder: " & csvFolderPath
End Sub
