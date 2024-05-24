#!/usr/bin/env python
# coding: utf-8


import tkinter as tk
from PIL import ImageGrab
import pytesseract
import pyperclip
import sys
from pynput import keyboard
from pynput.mouse import Listener as MouseListener
from pynput.keyboard import Key, Listener as KeyboardListener

class CaptureOnClick :
    def __init__(self) :
        self.root = tk.Tk()
        self.root.attributes('-fullscreen', True)
        self.root.attributes('-alpha', 0.2)
        self.root.configure(bg='red')
        self.root.bind('<Escape>', lambda e : e.widget.quit())
        
        self.canvas = tk.Canvas(self.root, cursor='cross')
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        self.start_position = None
        self.end_position = None
        
        self.canvas.bind('<ButtonPress-1>', self.on_click_start)
        self.canvas.bind('<B1-Motion>', self.on_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_click_end)
        
    def on_click_start(self, event) :        
        self.start_position = (event.x, event.y)
        self.canvas.delete('rect')
        
    def on_drag(self, event) :
        self.end_position = (event.x, event.y)
        self.canvas.delete('rect')
        self.canvas.create_rectangle(self.start_position[0],
                                     self.start_position[1],
                                     self.end_position[0],
                                     self.end_position[1], outline='red',
                                     width=3,
                                     tags='rect')
        
    def on_click_end(self, event) :
        self.end_position = (event.x, event.y)
        self.capture()
        self.root.quit()
        
    def capture(self) :
        if None not in (self.start_position, self.end_position) :
            x1, y1 = self.start_position
            x2, y2 = self.end_position
            bbox = (x1, y1, x2, y2)
            
            x1 += self.root.winfo_x()
            y1 += self.root.winfo_y()
            x2 += self.root.winfo_x()
            y2 += self.root.winfo_y()
            
            self.root.withdraw()
            screenshot = ImageGrab.grab(bbox)
            screenshot.convert('L')
            screenshot.save("capture_img.png")
            
            text = pytesseract.image_to_string(screenshot, lang='eng').strip()
            pyperclip.copy(text)
            print(text)
            
    def run(self) :
        self.root.mainloop()
        
def on_press(key) :
    try :
        if key== keyboard.Key.ctrl_l or key == keyboard.Key_ctrl_r :
            global ctrl_pressed
            ctrl_pressed = True
    
    except AttributeError :
            pass
            
def on_release(key) :
    if key == keyboard.Key.esc :
        is_exit = input("종료(y/n)?")
        if is_exit == 'y' or is_exit == 'Y' :
            sys.exit()
        
    try :
        global ctrl_pressed
        if (key == keyboard.Key.ctrl_l or key == keyboard.Key.ctrl_r) and ctrl_pressed :
            ctrl_pressed = False
    
    except AttributeError :
        pass
    
def on_click(x, y, button, pressed) :
    if button.name == 'left' and pressed and ctrl_pressed :
        cap = CaptureOnClick()
        cap.run()
        

if __name__ == '__main__' :
    print("""
    ############# Claude Monet 사용 가이드 ####################
    1. Tesseract 엔진 설치 및 사용자 환경 변수 경로 설정 필수
    2. (컨트롤 + 마우스 왼쪽 버튼)으로 캡처 모드로 진입
    3. 캡처 모드에서는 마우스 왼쪽 누른 상태로 드래그하여 캡처 영역 지정
    4. 캡처된 이미지 파일을 가지고 tesseract 엔진 ocr 작업을 실시하여 이미지를 텍스트로 반환 후 콘솔창에 띄우고 클립보드에도 저장하는 프로그램임
    5. 클립보드에 저장되어 있으므로 바로 (컨트롤+v)로 붙여넣기 가능
    6. 한글 OCR 지원 안함(성능 좋지 않음. 영어까지 퀄이 낮아짐)
    7. 표 형태로 문단 구분하여 캡처 가능. 단 너무 많은 영역의 경우 속도가 느려짐
    8. 캡처 모드에서 빠져나오는 단축키는 없음. 그냥 아무 영역이나 한번 긁어서 창을 종료시키던가, 작업 표시줄에 띄워진 깃털팬 모양의 프로그램을 수동으로 닫아야 함(첫번째 방법 추천)
    9. 값이 콘솔창에 뜨지 않으면 캡처한 부분을 인식하지 못하는 것임
    10. 드래그 상자 및 캡처 모드 흐릿하게 나오는거 바꿔보려 했지만 잘 안됨
    11. 메인 모니터 창 외에 다른 창은 캡처 불가능
    12. 같은 이미지라 할지라도 얼마만큼의 배율로 어떻게 긁었느냐에 따라 인식 성능이 달라짐
    13. 글자가 이상하게 인식되는건 프로그램 문제가 아니라 ocr 엔진 문제로 해결해 줄 수 없음
    14. 기능 추가 및 업그레이드 계획 없음. 필요시 소스코드는 공유해 줄 수 있음
    15. 프로그램이 실행되는 콘솔 창의 값은 캡처할 수 없음
    16. GS 프로젝트 외 사용이 필요할 경우 코드를 다시 타이핑 해야 하므로, 별도 논의 필요
    17. 버그 제보 : 010-5096-4025
    -----------------------------------------------------------------""")
    
    ctrl_pressed = False
    
    with KeyboardListener(on_press=on_press, on_release=on_release) as k_listener :
        with MouseListener(on_click=on_click) as m_listener :
            k_listener.join()
            m_listener.join()        

