#Persistent  ; 스크립트를 지속적으로 실행
#NoEnv  ; 환경변수 사용 안 함
#Warn  ; 경고 활성화
#SingleInstance force  ; 중복 실행 방지
SendMode Input  ; 입력 모드 설정

stopping := false

;; GUI 프로그램 시작 화면 창 구성 -----------------------------------------------------------------------------------

Gui, +LastFound
Gui, +MinimizeBox  ; 최소화 버튼 활성화

hGui := WinExist()

manual =
(
무엇이든 상상할 수 있는 사람은
무엇이든 만들어 낼 수 있다
				- 엘런 튜링

단축키 목록

■ ALT + Q : 열려 있는 엑셀 창 활성화
■ ALT + CapsLock : EVERYTHING 활성화 시 자동 검색 및 파일 열기
■ ALT + D : CAD 도면 확대 후 좌하단으로 이동
■ ALT + P : 일시정지/일시정지 해제
■ F12 : 종료

                                                2024.08.27 version 1.3.0

)

Gui, Add, Text, , %manual%
Gui, Add, Button, x156 y170 w70 h40 gMinimizeBtn Default, 확인  ; 최소화 버튼 추가. 확인 버튼에 기본 포커스
Gui, Add, Button, x249 y170 w70 h40 gExitBtn, 종료
Gui, Show, w360 h220, 꼬북이A

Menu, Tray, Add, 첫 화면, RestoreGui
Menu, Tray, Add, 종료 (F12), ExitScript
Menu, Tray, Default, 첫 화면
Menu, Tray, Tip, 꼬북이  ; 트레이 아이콘에 툴팁 설정

return

MinimizeBtn:  ; 최소화 버튼 라벨
    WinHide, ahk_id %hGui%
return

RestoreGui: ; 첫 화면 라벨
    WinShow, ahk_id %hGui%
    WinActivate, ahk_id %hGui%
	SendInput {Esc}
return

ExitScript:
    ExitApp
return

ExitBtn:
    ExitApp  ; 종료 버튼 클릭 시 종료
return

GuiClose:
    ExitApp  ; GUI 창 닫기 버튼 클릭 시 종료
return


;; HOT KEY 명령 설정 -----------------------------------------------------------------------------------

!q::  ; Alt+Q를 누를 때 실행
IfWinExist, ahk_class XLMAIN  ; Excel 창이 존재하는지 확인
    WinActivate  ; 존재한다면 활성화
    SendInput {Esc}
return

!CapsLock:: ; Everything을 바로 활성화
    SendInput ^c
    sleep 300
IfWinExist, ahk_class EVERYTHING
    WinActivate
    SendInput {Tab}
    SendInput ^v
    SendInput {Tab}
    SendInput {Enter}

return


!d::
    MouseClick, WheelUp, , , 3  ; 마우스 휠을 5번 위로 스크롤
    Sleep, 100  ; 100밀리초 동안 대기
    Click Down Middle  ; 마우스 가운데 버튼을 누름 (Down)
    MouseMove, 100, -300, 100, R  ; 현재 위치에서 상대적으로 우상단(200, -200)으로 100밀리초 동안 이동
    Click Up Middle  ; 마우스 가운데 버튼을 놓음 (Up))
    MouseMove, -100, 300, 100, R ;
return


; 일시정지/재시작
!p::
    stopping := !stopping
    if (stopping) {
        Hotkey, !q, Off
        Hotkey, !CapsLock, Off
        Hotkey, !d, Off
        Tooltip, 매크로 비활성화
    }else {
        Hotkey, !q, On
        Hotkey, !CapsLock, On
        Hotkey, !d, On
        Tooltip, 매크로 활성화
    }
    Sleep, 1000
    Tooltip
return


; 종료
F12::
	MsgBox, 종료되었습니다
	ExitApp
