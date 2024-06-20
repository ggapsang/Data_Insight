Sub MatchKeysInRangNumber()
'''
''' 복합 키에서 한쪽 key 값이 텍스트나 단일한 정수로 매핑 테이블에서 찾을 수 있지만,
''' 다른 나머지 값은 매핑 테이블에 min~max 와 같은 범위로 되어 있는 수 일때
''' 매핑 테이블에서 키 값을 찾아 새로운 시트에 기록한다
''' 배관 PMS 테이블 매핑 시 사용할 수 있다


'기본 변수 선언
    
        
    ' 워크시트 변수 설정
    Dim ws1 As Worksheet 'pms 테이블이 있는 워크시트
    Dim ws2 As Worksheet 'pms 테이블 값을 가지고 올 개별속성리스트 워크시트
    
    '배열 선언을 위해 마지막 행 검색
    Dim lastRowS1 As Long 'ws1의 마지막 행. KEY가 있는 A열을 가지고 검색
    Dim lastRowS2 As Long 'ws2의 마지막 행, 정확한 행 카운팅을 위해 추후 'SR No'가 있는 행을 열을 정의할 것
        'S2시트의 경우 해더 행 역시 동적 설정이 필요함(사용자에게 input으로 받을 것)
    
    
    '배열에서 key 값을 찾기 위해 선언하는 매직 넘버
    Dim i As Long, j As Long
    
    
    'ws1의 변수들
    Dim keyS1 As String 'ws1의 key값
    Dim pmsCodeS1 As String 'ws1의 pms code 값
    
    Dim minS1 As Double 'ws1의 min(float) 값
    Dim maxS1 As Double 'ws1의 max(float) 값
    
    
    'ws2의 변수들
    Dim keyS2 As String 'ws2의 key값
    Dim pmsCodeS2 As String 'ws2의 pms code 값
    Dim sizeS2 As Double 'ws2의 사이즈 값
    
    '매칭 여부에 대한 플래그
    Dim matchFound As Boolean
    

'시트 정의
    
    Set ws1 = ThisWorkbook.Sheets("확정_여기서 작업")
    Set ws2 = ActiveWorkbook.Sheets("개별속성 배관") '파일 불러오기로 동적으로 정의될 수 있음
    
    
    ' 마지막 행 찾기
    lastRowS1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRowS2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    
    
    
'데이터를 배열에 저장
    
    'ws1의 데이터를 배열에 저장
    
    Dim keyArrS1 As Variant 'key를 저장할 배열
    Dim codeArrS1 As Variant 'pms code를 저장할 배열
    Dim minArrS1 As Variant 'min 값을 저장할 배열
    Dim maxArrS1 As Variant 'max 값을 저장할 배열
    
    keyArrS1 = ws1.Range("A2:A" & lastRowS1).value
    codeArrS1 = ws1.Range("B2:B" & lastRowS1).value
    minArrS1 = ws1.Range("C2:C" & lastRowS1).value
    maxArrS1 = ws1.Range("D2:D" & lastRowS1).value
    
    
    'ws2의 데이터를 배열에 저장
    
    Dim keyArrS2 As Variant 'key를 저장할 배열
    Dim codeArrS2 As Variant 'pms code를 저장할 배열
    Dim sizeArrS2 As Variant 'size를 저장할 배열
    
    
    keyArrS2 = ws2.Range("A2:A" & lastRowS2).value '이후에 동적으로 설정되어야 하는 칼럼 번호와 행 번호
    codeArrS2 = ws2.Range("B2:B" & lastRowS2).value '이후에 동적으로 설정되어야 하는 칼럼 번호와 행 번호
    sizeArrS2 ws2.Range("C2:C" & lastRowS2).value '이후에 동적으로 설정되어야 하는 칼럼 번호와 행 번호
    
    
    
'두 개의 시트의 비교 및 결과 저장

    For i = 1 To UBound(keyArrS2, 1)
        
        key2 = keyArrS2(i, 1)
        code2 = codeArrS2(i, 1)
        size2 = sizeArrS2(i, 1)
        matchFound = False
            
            
        For j = 1 To UBound(keyArrS1, 1)
            
            key1 = keyArrS1(j, 1)
            code1 = codeArrS1(j, 1)
            min1 = minArrS1(j, 1)
            max1 = maxArrS1(j, 1)
            
            
            If code1 = code2 And size2 >= min1 And size2 <= max1 Then
                
                '여기에 수행 사항을 기록한다(별도 서브 프로시저를 만드는 것을 권장)
                
                matchFound = True
                Exit For
                
            End If
        
        Next j
    
            If Not matchFound Then
            
                '여기에 수행 사항을 기록한다(별도 서브프로시저를 만드는 것을 권장)
            
            
            End If
            
    Next i

    MsgBox "완료"

End Sub


Public Sub ImportPmsValue()




End Sub


Public Sub NotifyNotMatch()

End Sub
