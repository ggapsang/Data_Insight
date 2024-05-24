Sub ConsolidateSelectedSheets()
    Dim ws As Worksheet
    Dim masterWs As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim startRow As Long
    Dim sheetName As String
    Dim selectedSheets As Collection
    Dim wsName As Variant
    Dim response As Integer

    ' 마스터 시트 생성
    Set masterWs = ThisWorkbook.Sheets.Add
    masterWs.Name = "ConsolidatedData"

    ' 선택된 시트 이름을 저장할 컬렉션 생성
    Set selectedSheets = New Collection

    ' 사용자가 통합할 시트를 선택하도록 메시지 박스를 표시
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> masterWs.Name Then
            response = MsgBox("시트를 통합하시겠습니까? " & ws.Name, vbYesNo)
            If response = vbYes Then
                selectedSheets.Add ws.Name
            End If
        End If
    Next ws

    ' 선택된 시트들을 통합
    startRow = 1
    For Each wsName In selectedSheets
        Set ws = ThisWorkbook.Sheets(wsName)
        With ws
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column

            ' 데이터를 마스터 시트로 복사
            .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Copy masterWs.Cells(startRow, 1)
            startRow = startRow + lastRow
        End With
    Next wsName

    MsgBox "선택한 시트 통합 완료!"
End Sub
