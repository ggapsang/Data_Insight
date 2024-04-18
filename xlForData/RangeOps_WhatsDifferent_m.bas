Sub WhatsDifferent()

    Dim rngHeader_1 As Range
    Dim rngHeader_2 As Range
    Dim rngKey_1 As Range
    Dim rngKey_2 As Range
    Dim inputRange As Range
    Dim isSkipEmpty As Boolean

''' 1. 사용자 입력 받기

    Set inputRange = Application.InputBox("첫번째 key 값 범위 선택 :", Type:=8)
    Set rngKey_1 = inputRange

    Set inputRange = Application.InputBox("첫번쨰 column 값 범위 선택 : ", Type:=8)
    Set rngHeader_1 = inputRange

    Set inputRange = Application.InputBox("두번째 key 값 범위 선택 : ", Type:=8)
    Set rngKey_2 = inputRange

    Set inputRange = Application.InputBox("두번째 column 값 범위 선택 : ", Type:=8)
    Set rngHeader_2 = inputRange

    isSkipEmpty = Application.InputBox(prompt:="빈 셀 생략(true=1, false=0) : ", Type:=4)

''' 2. 각각의 값을 배열에 저장

    ''첫 번째 시트의 값들을 배열에 저장

    Dim keyArray_1() As Variant '키 배열
    Dim colArray_1() As Variant '열 배열
    Dim valueArray_1() As Variant '값 배열
    
    keyArray_1() = rngKey_1.Value
    colArray_1() = rngHeader_1.Value

    '시작 행/열, 끝 행/열 저장
    Dim beginRow As Long, endRow As Long, beginCol As Integer, endCol As Integer
    beginRow = rngKey_1.Row
    endRow = beginRow + rngKey_1.Rows.Count - 1
    beginCol = rngHeader_1.Column
    endCol = beginCol + rngHeader_1.Count - 1

    Dim ws_1 As Worksheet, wb_1 As Workbook '워크북과 시트 이름
    Set ws_1 = rngKey_1.Parent
    Set wb_1 = ws_1.Parent

    valueArray_1 = ws_1.Range(ws_1.Cells(beginRow, beginCol), ws_1.Cells(endRow, endCol)).Value

    
    ''두 번째 시트에 기본 변수 선언 : 첫 번째 시트와 동일한 방식으로 저장
    
    Dim keyArray_2() As Variant '키 배열
    Dim colArray_2() As Variant '열 배열
    Dim valueArray_2() As Variant '값 배열
    
    keyArray_2() = rngKey_2.Value
    colArray_2() = rngHeader_2.Value

    '시작 행/열, 끝 행/열 저장
    beginRow = rngKey_2.Row
    endRow = beginRow + rngKey_2.Rows.Count - 1
    beginCol = rngHeader_2.Column
    endCol = beginCol + rngHeader_2.Count - 1

    Dim ws_2 As Worksheet, wb_2 As Workbook '워크북과 시트 이름
    Set ws_2 = rngKey_2.Parent
    Set wb_2 = ws_2.Parent

    valueArray_2 = ws_2.Range(ws_2.Cells(beginRow, beginCol), ws_2.Cells(endRow, endCol)).Value

''' 3. 새 통합 문서 생성 후 해당 문서에 서로 다른 입력 리스트들을 정리함
    Dim compareWb As Workbook
    Dim compareWs As Worksheet
    
    Set compareWb = Workbooks.Add
    Set compareWs = compareWb.Worksheets(1)
    
    compareWs.Range("A1").Value = "Key"
    compareWs.Range("B1").Value = "Column"
    compareWs.Range("C1").Value = "value_1"
    compareWs.Range("D1").Value = "value_2"
    
    Dim workRow As Long
    workRow = 2
    
    
''' 4. 딕셔너리를 생성하고 key-col을 기준으로 매핑하면서 배열의 값 찾기

    Dim dictKeyMap As Object
    Dim dictColMap As Object
    
    Set dictKeyMap = CreateObject("Scripting.Dictionary")
    Set dictColMap = CreateObject("Scripting.Dictionary")
    
    '' keyArray_1를 딕셔너리에 매핑
    For k1 = LBound(keyArray_1, 1) To UBound(keyArray_1, 1)
        dictKeyMap(keyArray_1(k1, 1)) = k1
    Next k1


    '' colArray_1의 칼럼 이름을 딕셔너리에 매핑
    For col_1 = LBound(colArray_1, 2) To UBound(colArray_1, 2)
        dictColMap(colArray_1(1, col_1)) = col_1
    Next col_1


    '' keyArray_2를 순회하며 일치하는 키 찾기
    For k2 = LBound(keyArray_2, 1) To UBound(keyArray_2, 1)
        
        Dim i_1 As Long
        Dim j_1 As Long
        Dim j_2 As Long
        
        If dictKeyMap.Exists(keyArray_2(k2, 1)) Then
            i_1 = dictKeyMap(keyArray_2(k2, 1))
            
            '일치하는 칼럼 값 찾기
            For col_2 = LBound(colArray_2, 2) To UBound(colArray_2, 2)
                If dictColMap.Exists(colArray_2(1, col_2)) Then
                
                    j_1 = dictColMap(colArray_2(1, col_2))
                    j_2 = col_2
            
            
                    If valueArray_1(i_1, j_1) <> valueArray_2(k2, col_2) Then
                        
                        ' isSkipEmpty가 True이고 둘 다 비어있지 않을 때, 혹은 isSkipEmpty가 False일 때
                        If (isSkipEmpty And (valueArray_1(i_1, j_1) <> "" And valueArray_2(k2, col_2) <> "")) Or Not isSkipEmpty Then
                        
                        compareWs.Range("A" & workRow).Value = keyArray_2(k2, 1)
                        compareWs.Range("B" & workRow).Value = colArray_2(1, col_2)
                        compareWs.Range("C" & workRow).Value = valueArray_1(i_1, j_1)
                        compareWs.Range("D" & workRow).Value = valueArray_2(k2, col_2)
        
                        workRow = workRow + 1
                        
                        End If
                    End If
                End If
            Next col_2
        End If
    Next k2


End Sub

