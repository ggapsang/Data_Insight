Sub SaveAllSheetsAsCSV()
    Dim ws As Worksheet
    Dim csvFolderPath As String
    Dim wb As Workbook
    Dim folderName As String
    Dim fso As Object
    Dim fileName As String
    
    Dim totalSheet As Integer
    Dim currentSheet As Integer
    Dim sheetNames() As String
    Dim i As Integer
    
    ' FileSystemObject 생성
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 현재 워크북의 파일 이름을 확장자 없이 가져옴
    fileName = ActiveWorkbook.Name
    folderName = fso.GetBaseName(fileName)
    
    ' 시트 이름을 배열에 저장
    sheetNames = ListSheetNamesToArray()
    
    ' 총 시트 개수 계산
    totalSheet = UBound(sheetNames) - LBound(sheetNames) + 1
    
    ' 폴더 경로 설정
    csvFolderPath = ActiveWorkbook.Path & "\" & folderName & "\"
    
    ' 폴더가 존재하지 않으면 생성
    If Dir(csvFolderPath, vbDirectory) = "" Then
        MkDir csvFolderPath
    End If
    
    ' 배열의 시트 이름들만 반복
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = ActiveWorkbook.Sheets(sheetNames(i))
        currentSheet = i - LBound(sheetNames) + 1
        
        ' 진행 상황 표시
        Application.StatusBar = "Processing sheet " & currentSheet & " of " & totalSheet & ": " & ws.Name
        
        ' CSV 파일로 저장
        SaveSheetAsCSV ws, csvFolderPath & ws.Name & ".csv"
    Next i
    
    ' 상태 표시줄 초기화
    Application.StatusBar = False
    
    MsgBox "All sheets have been saved as CSV files in folder: " & csvFolderPath
End Sub

Function ListSheetNamesToArray() As String()
    Dim ws As Worksheet
    Dim i As Integer
    Dim idx_sheet_nm As String
    Dim sheetNames() As String
    
    idx_sheet_nm = "Sheet_Name_list"
    
    ' Initialize the array to hold sheet names with an initial size
    ReDim sheetNames(1 To 1)
    
    i = 0
    For Each ws In Worksheets
        ' Store sheet names in the array
        If ws.Name <> idx_sheet_nm Then
            i = i + 1
            ReDim Preserve sheetNames(1 To i)
            sheetNames(i) = ws.Name
        End If
    Next ws
    
    ListSheetNamesToArray = sheetNames
End Function

Sub SaveSheetAsCSV(ws As Worksheet, filePath As String)
    Dim fs As Object
    Dim aCell As Range
    Dim rowNum As Long
    Dim colNum As Long
    Dim csvLine As String
    Dim replaceWith As String
    
    replaceWith = ";"
    
    Set fs = CreateObject("Scripting.FileSystemObject").CreateTextFile(filePath, True)
    
    Application.DisplayAlerts = False
    
    For rowNum = 1 To ws.UsedRange.Rows.Count
        csvLine = ""
        For colNum = 1 To ws.UsedRange.Columns.Count
            Set aCell = ws.Cells(rowNum, colNum)
            ' 모든 셀 값을 문자열로 처리하고 따옴표로 감쌈
            csvLine = csvLine & """" & Replace(Replace(CStr(aCell.Value), vbLf, replaceWith), vbCr, replaceWith) & """" & ","
        Next colNum
        csvLine = Left(csvLine, Len(csvLine) - 1)
        fs.WriteLine csvLine
    Next rowNum
    
    fs.Close
    Application.DisplayAlerts = True
End Sub
