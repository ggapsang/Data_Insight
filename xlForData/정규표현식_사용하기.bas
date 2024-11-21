Sub ReplaceByRe()
    Dim ws As Worksheet
    Dim cell As Range
    Dim regex As Object
    Dim matches As Object
    Dim inputText As String
    
    ' 정규식 객체 생성
    Set regex = CreateObject("VBScript.RegExp")
    '예시 regex.Pattern = "FROM\s+\w+\s" ' "FROM 공백 단어 공백" 패턴
    regex.Pattern = input()
    regex.Global = True
    
    ' 현재 활성화된 시트를 대상으로 작업
    Set ws = ActiveWorksheet
    
    ' 모든 셀을 반복
    For Each cell In ws.Range("c2:c2201")
        If Not IsEmpty(cell.Value) Then
            inputText = cell.Value
            ' 패턴이 매칭되면 교체
            If regex.test(inputText) Then
                cell.Value = regex.Replace(inputText, "FROM table")
            End If
        End If
    Next cell
    
    MsgBox "변경이 완료되었습니다!"
End Sub

