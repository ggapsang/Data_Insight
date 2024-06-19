Sub SaveAllSheetsAsCSV()
    Dim ws As Worksheet
    Dim csvFolderPath As String
    Dim wb As Workbook
    Dim tempWb As Workbook
    Dim folderName As String
    Dim fso As Object
    Dim fileName As String
    Dim tempWs As Worksheet
    
    Dim totalSheet As Integer
    Dim currentSheet As Integer
    
    ' FileSystemObject 생성
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 현재 워크북의 파일 이름을 확장자 없이 가져옴
    fileName = ActiveWorkbook.Name
    folderName = fso.GetBaseName(fileName)
    
    ' 시트 이름을 배열에 저장
    Dim sheetNames() As String
    sheetNames = ListSheetNamesToArray()
    
    ' 폴더 경로 설정
    csvFolderPath = ThisWorkbook.Path & "\" & folderName & "\"
    
    ' 폴더가 존재하지 않으면 생성
    If Dir(csvFolderPath, vbDirectory) = "" Then
        MkDir csvFolderPath
    End If
    
    ' 배열의 시트 이름들만 반복
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        currentSheet = currentSheet + 1
        
        ' 임시 워크북 생성
        ws.Copy
        
        Set tempWb = ActiveWorkbook
        Set tempWs = tempWb.Sheets(1)
             
        tempWs.Range("A1").CurrentRegion.Select
        
        BreakandFill2 Selection
             
        ' CSV 파일로 저장 (기존 파일 덮어쓰기)
        Application.DisplayAlerts = False
        tempWb.SaveAs fileName:=csvFolderPath & ws.Name & ".csv", FileFormat:=xlCSV, CreateBackup:=False
        Application.DisplayAlerts = True
        
        ' 임시 워크북 닫기
        tempWb.Close SaveChanges:=False
        
    Next i
    
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

Public Sub BreakandFill2(rngWork As Range)
    Dim header_row As Integer
    Dim cardNo_col As Integer
    Dim cardNo_col_chr As String
    Dim cell As Range
    
    Dim replaceWith As String
    
    replaceWith = ";"
    
    Application.DisplayAlerts = False

    For Each cell In rngWork
        Dim value As Variant
        value = cell.value
        If InStr(1, value, vbLf) > 0 Then
            value = Replace(value, vbLf, replaceWith)
            cell.value = value
        ElseIf InStr(1, value, vbCr) > 0 Then
            value = Replace(value, vbCr, replaceWith)
            cell.value = value
        End If
    Next cell

    For Each cell In rngWork
        If cell.MergeCells Then
            Dim mergeRange As Range
            Set mergeRange = cell.MergeArea
            cell.UnMerge
            value = cell.value
            value = Replace(value, vbCrLf, replaceWith)
            mergeRange.value = value
        End If
    Next cell
    
    Application.DisplayAlerts = True
End Sub


