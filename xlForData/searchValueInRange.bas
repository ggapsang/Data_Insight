' 함수: FindValueInRange
' 설명: 주어진 코드와 크기를 기준으로 지정된 범위에서 값을 검색
'       코드가 일치하는지, 크기는 최소 및 최대 범위 내에 있는지 확인
' 매개변수:
'   - code (String): keyRange에서 일치하는 코드 찾기
'   - size (Single): minRange와 maxRange 내에 있는지 확인할 크기
'   - keyRange (Range): 코드(key)를 포함하는 범위
'   - minRange (Range): 각 코드의 최소값을 포함하는 범위
'   - maxRange (Range): 각 코드의 최대값을 포함하는 범위
'   - returnRange (Range): 각 코드에 해당하는 반환 값을 포함하는 범위
' 반환값:
'   - 일치하는 값이 있는 경우 returnRange에서 해당 값을 반환. 그렇지 않으면 #N/A 오류를 반환

Function FindValueInRange(code As String, size As Single, keyRange As Range, minRange As Range, maxRange As Range, returnRange As Range) As Variant

    Dim keys() As Variant
    Dim mins() As Variant
    Dim maxs() As Variant
    Dim values() As Variant

    keys = keyRange.Value
    mins = minRange.Value
    maxs = maxRange.Value
    values = returnRange.Value

    Dim i As Integer
    
    For i = LBound(keys, 1) To UBound(keys, 1)
        If keys(i, 1) = code Then
            If mins(i, 1) <= size And maxs(i, 1) >= size Then
                FindValueInRange = values(i, 1)
                Exit Function
            End If
        End If
    Next i

    FindValueInRange = CVErr(xlErrNA)
    
End Function
