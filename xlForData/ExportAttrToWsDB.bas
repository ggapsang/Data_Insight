Sub ExportWsDB2()

''' 작업장 시트 B열에 속성 값이 없을 경우 오버플로우 오류 발생함

    Application.ScreenUpdating = False '화면상태 표시 비활성화
    Application.Calculation = xlManual '자동 계산 비활성화
    Application.DisplayStatusBar = False '상태 표시줄 비활성화
    Application.EnableEvents = False '이벤트 비활성화


' wsDB sheet
'' 'A1' SR No : key Value
'' 'O1 : FV1' : individual value_cols

' wsWork Sheet
'' 'C5 : end(xltoright) : SR No - key Value
'' 'A13 : end(xltoDown) : individual value_cols


'워크시트 변수

    Dim wsDB As Worksheet
    Dim wsWork As Worksheet
    
    Set wsDB = ThisWorkbook.Sheets("개별속성리스트")
    Set wsWork = ThisWorkbook.Sheets("개별속성리스트_작업장")

    Dim rngSrNoDB As Range
    Dim rngSrNoWork As Range
    Dim rngIndivValueDB As Range
    Dim rngIndivValueWork As Range
    
    Dim strSrNoDB As String
    Dim strSrNoWork As String
    Dim strIndividualRow As String
    Dim j As Integer

    strSrNoDB = "A2" 'wsDB에서 SR No가 시작되는 좌표(내부 해더 포함)
    strSrNoWork = "C5" 'wsWork에서SR No가 시작되는 좌표
    
'''' 포멧 수정시 주목
    strIndividualRow = "B19"
    j = 18 'strRowindividual 변수의 행 번호 -1
    
    Set rngSrNoDB = wsDB.Range(strSrNoDB, wsDB.Range(strSrNoDB).End(xlDown)) 'SR No가 시작되는 열 번호
    Set rngSrNoWork = wsWork.Range(strSrNoWork, wsWork.Range(strSrNoWork).End(xlToRight)) 'SR No가 있는 좌표
    
    Dim intColIndiv_1 As Integer
    intColIndiv_1 = wsDB.Range("A1", wsDB.Range("A1").End(xlToRight)).Find("속성1", LookIn:=xlValues).Column
        
        
    For Each rngSr In rngSrNoWork
    
        
        Dim intRowNo As Integer
        Dim intColNo As Integer
        
        intRowNo = rngSrNoDB.Find(rngSr.value, LookIn:=xlValues).Row
        intColNo = rngSr.Column
        strColLetter = ConvertToLetter(intColNo)
        
        Dim intCountIndiv As Integer
        
        On Error Resume Next
        intCountIndiv = wsWork.Range(strIndividualRow).End(xlDown).Row
        
        If Err.Number <> 0 Then
            intCountIndiv = wsWork.Range(strIndividualRow).Row
            Err.Clear
        End If
        On Error GoTo 0
        
        Dim rngIndiviValueWork As Range
        Set rngIndiviValueWork = wsWork.Range(strColLetter & j + 1 & ":" & strColLetter & intCountIndiv)
        
        Dim i As Integer
        i = 0
        For Each rngIndiv In rngIndiviValueWork
            wsDB.Cells(intRowNo, i + intColIndiv_1).value = rngIndiv.value
            i = i + 1
        Next
        
    Next
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True '상태 표시줄 비활성화
    Application.EnableEvents = True '이벤트 비활성화

End Sub

Function ConvertToLetter(iCol As Integer) As String
   Dim a As Integer
   Dim b As Integer
   a = iCol
   ConvertToLetter = ""
   Do While iCol > 0
      a = Int((iCol - 1) / 26)
      b = (iCol - 1) Mod 26
      ConvertToLetter = Chr(b + 65) & ConvertToLetter
      iCol = a
   Loop
End Function

Function FindValueInRow(searchRow As Long, searchValue As Variant, ws As Worksheet) As String

    ' Define variables
    Dim rng As Range
    Dim foundCell As Range
    Dim colLetter As String
    
    ' Set the range to search within
    Set rng = ws.Rows(searchRow)
    
    ' Search for the value
    Set foundCell = rng.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' If the value is found, get the column letter
    If Not foundCell Is Nothing Then
        colLetter = Replace(foundCell.Cells.Address(False, False), searchRow & "", "")
        FindValueInRow = colLetter
    Else
        FindValueInRow = "Value not found."
    End If

End Function


Sub export2()

    ExportWsDB2

End Sub

