Sub CompareValues()

    Dim rngHeader_1 as Range
    Dim rngHeader_2 as Range
    Dim rngKey_1 as Range
    Dim rngKey_2 as Range
    Dim inputRange as Range
    Dim isSkipEmpty as Boolean

''' 1. 사용자 입력 받기

    Set inputRange = Application.InputBox("첫번째 key 값 범위 선택 :", Type:=8)
    Set rngHeader_1 = inputRange

    Set inputRange = Application.InputBox("첫번쨰 column 값 범위 선택 : ", Type:=8)
    Set rngKey_1 = inputRange

    Set inputRange = Application.InputBox("두번째 key 값 범위 선택 : ", Type:=8)
    Set rngHeader_2 = inputRange

    Set inputRange = Application.InputBox("두번째 column 값 범위 선택 : ", Type:=8)
    Set rngHeader_2 = inputRange

    isSkipEmpty = Application.InputBox(prompt:="빈 셀 생략(true=1, false=0) : ", Type:=4)

''' 2. 각각의 값을 배열에 저장

    ''첫 번째 시트의 값들을 배열에 저장

    Dim keyArray_1() As Variant '키 배열
    Dim colArray_1() As Variant '열 배열
    Dim valueArray_1() As Variant '값 배열
    
    keyArray_1() = rngKey_1.value
    colArray_1() = rngHeader_1.value

    '시작 행/열, 끝 행/열 저장
    Dim beginRow as Long, endRow as Long, beginCol as Integer, endCol as Integer
    beginRow = rngKey_1.Row
    endRow = beginRow + rngKey_1.Rows.Count - 1
    beginCol = rngHeader_1.Column
    endCol = endCol + rngHeader_1.Count - 1

    Dim ws_1 as Worksheet, wb_1 as Workbook '워크북과 시트 이름
    Set ws_1 = rngKey_1.Parent 
    Set wb_1 = wb_1.Parent

    valueArray_1() = ws_1.range(ws_1.Cells(beginRow, beginCol), ws_1.Cells(endRow, endCol)).Value
    
    ''두 번째 시트에 기본 변수 선언 : 첫 번째 시트와 동일한 방식으로 저장
    
    Dim keyArray_2() As Variant '키 배열
    Dim colArray_2() As Variant '열 배열
    Dim valueArray_2() As Variant '값 배열
    
    keyArray_2() = rngKey_2.value
    colArray_2() = rngHeader_2.value

    '시작 행/열, 끝 행/열 저장
    Dim beginRow as Long, endRow as Long, beginCol as Integer, endCol as Integer
    beginRow = rngKey_2.Row
    endRow = beginRow + rngKey_2.Rows.Count - 1
    beginCol = rngHeader_2.Column
    endCol = endCol + rngHeader_2.Count - 1

    Dim ws_2 as Worksheet, wb_2 as Workbook '워크북과 시트 이름
    Set ws_2 = rngKey_2.Parent 
    Set wb_2 = wb_2.Parent

    valueArray_2() = ws_2.range(ws_2.Cells(beginRow, beginCol), ws_2.Cells(endRow, endCol)).Value

''' 3. 새 통합 문서 생성 후 해당 문서에 서로 다른 입력 리스트들을 정리함
    Dim compareBook as Workbook
    Dim compareSheet as Wokrsheet

''' 4. 딕셔너리를 생성하고 key-col을 기준으로 매핑하면서 배열의 값 찾기

End sub
