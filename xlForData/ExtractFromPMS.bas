Public Sub process_single_worksheet(ws As Worksheet)

'a1에 대고 break and fill 매서드 실행

    Dim rngWork As Range

    Range("A1").CurrentRegion.Select
    
    BreakandFill Selection


'필터링을 위한 기본 변수 설정

    Dim rngShortcodeCol As Range
    Dim rngMinCol As Range
    Dim rngMaxCol As Range
    Dim rngDescriptCol As Range
    
    
    Dim intShortcodeColidx As Integer
    Dim intMinColidx As Integer
    Dim intMaxColidx As Integer
    Dim intDescriptColidx As Integer
    
    
    Dim destCell As Range
    Dim cell As Range
    
    
    
    Dim intFilterRow As Integer
    
    intFilterRow = 9
    
    Set rngShortcodeCol = ws.Rows(intFilterRow).Find(what:="SHORT CODE", LookIn:=xlValues, LookAt:=xlWhole)
    Set rngMinCol = ws.Rows(intFilterRow).Find(what:="MIN.", LookIn:=xlValues, LookAt:=xlWhole)
    Set rngMaxCol = ws.Rows(intFilterRow).Find(what:="MAX.", LookIn:=xlValues, LookAt:=xlWhole)
    Set rngDescriptCol = ws.Rows(intFilterRow).Find(what:="DESCRIPTION", LookIn:=xlValues, LookAt:=xlWhole)
    
    intShortcodeColidx = rngShortcodeCol.Column
    intMinColidx = rngMinCol.Column
    intMaxColidx = rngMaxCol.Column
    intDescriptColidx = rngDescriptCol.Column
    
    
    Dim findColidx As Integer
    Dim findValue As String
    
    findColidx = intShortcodeColidx
    findValue = "P"
    
    
'복사 붙여넣기
    
    Dim concatWs As Worksheet
    Set concatWs = ThisWorkbook.Sheets("concat")
    
    
    Set destCell = concatWs.Range("O2") ''O2 : general material
    FilterAndPaste ws, intFilterRow, intShortcodeColidx, findValue, intDescriptColidx, destCell

    Set destCell = concatWs.Range("C2") ''c2 : min
    FilterAndPaste ws, intFilterRow, intMinColidx, findValue, intMinColidx, destCell

    Set destCell = concatWs.Range("D2") ''d2 : max
    FilterAndPaste ws, intFilterRow, intMaxColidx, findValue, intMaxColidx, destCell
    

' 다음 작업을 할 행 업데이트.
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    workRow = lastRow + 1

End Sub


Sub FilterAndPaste(ws As Worksheet, intFilterRow As Integer, findColidx As Integer, findValue As String, pasteColidx As Integer, destCell As Range)

    '필터 걸기
    ws.Rows(intFilterRow).Autofitler
    
    '특정 열의 인덱스에서 findColidx와 일치하는 항목 필터링
    ws.Rows(intFilterRow).AutoFilter Filed:=colidx, Criteria1:=findValue
    
    '붙여넣기
    For Each cell In ws.Range(ws.Cells(intFilterRow + 1, pasteColidx), ws.Cells(ws.Rows.Count, colidx)).SpecialCells(xlCellTypeVisible)
        
        destCell.value = cell.value
        Set destCell = destCell.Offset(1, 0)
    
    Next cell
    
    ws.AutoFilter = False


End Sub
