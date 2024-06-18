Sub ListAllSheetNames()

    Dim ws As Worksheet
    Dim i As Integer
    Dim idx_sheet_nm As String
    
    idx_sheet_nm = "Sheet_Name_list"
    
    '새 시트 생성
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = idx_sheet_nm
    i = 1
    For Each ws In Worksheets
    '시트 이름 리스트 업
        If ws.Name <> idx_sheet_nm Then
            Sheets(idx_sheet_nm).Cells(i, 1) = ws.Name
            i = i + 1
        End If
        
    Next ws
    
End Sub
