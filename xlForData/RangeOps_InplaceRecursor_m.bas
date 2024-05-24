Attribute VB_Name = "RangeOps_InplaceRecursor_m"
Sub InplaceRecursor()

    Dim msgResponse As Integer

    ' 순환 참조 설정 상태를 확인하여 적절한 메시지를 표시
    If Application.Iteration Then
        msgResponse = MsgBox("순환 참조를 비활성화하시겠습니까?", vbYesNo + vbQuestion, "순환 참조 설정 변경")
        If msgResponse = vbYes Then
            ' 순환 참조 비활성화
            Application.Iteration = False
            MsgBox "순환 참조가 비활성화되었습니다.", vbInformation
        End If
    Else
        msgResponse = MsgBox("순환 참조를 활성화하시겠습니까?", vbYesNo + vbQuestion, "순환 참조 설정 변경")
        If msgResponse = vbYes Then
            ' 순환 참조 활성화
            Application.Iteration = True
            Application.MaxIterations = 1  ' 순환 참조시 최대 반복 횟수를 1로 설정
            Application.MaxChange = 0.001  ' 허용되는 최대 변화량 설정
            MsgBox "순환 참조가 활성화되었습니다. 계산을 수행한 후에는 비활성화해주세요.", vbInformation
        End If
    End If

End Sub
