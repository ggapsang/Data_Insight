import fitz
import re
import os
import pandas as pd
import ezdxf

# 확장자를 제외한 파일명 추출
def get_file_name_strip(file) :
    basename = os.path.basename(file)
    file_name = os.path.splitext(basename)[0]
    
    return file_name

# 확장자 추출
def get_file_extension(file) :
    basename = os.path.basename(file)
    file_name = os.path.splitext(basename)[1]
    
    return file_name

# pdf 주석에서 텍스트 추출
def extract_annotations_from_pdf(pdf_path) :
    doc = fitz.open(pdf_path)
    annotations = []
    
    for page in doc :
        annots = page.annots()
        for annot in annots :
            info = annot.info
            annotations.append(info['content'] if 'content' in info else '')
                    
    doc.close()
    
    return annotations

# 정규표현식에서 알맞는 값 찾기
def find_texts(re_list, annot, remove_list) :
    
    candidate_texts = []
    for regression in re_list :
        for text in annot :
            if regression.findall(text) :
                candidate_texts.append(text)
    
    candidate_texts = list(set(candidate_texts))
    
    candidate_texts.sort(reverse=False)
    candidate_texts.sort(key=len, reverse=True)
    
    candidate_texts.sort()

    if candidate_texts != [] :
        longest_item = max(candidate_texts, key=len)
        candidate_texts.remove(longest_item)
        candidate_texts.insert(0, longest_item)
    
    return candidate_texts


# 후보 텍스트들에서 일부 제거

def preprocessing(candidate_texts, remove_list) :
    for char in remove_list :
        for text in candidate_texts :        
            if char in text :
                candidate_texts.remove(text)

    return candidate_texts


  # 리스트 길이 맞추기
def pad_list_to_length(original_list, target_length):
    # 리스트 길이가 목표 길이보다 작은 경우, 차이만큼 None을 추가
    while len(original_list) < target_length:
        original_list.append(None)
    return original_list

def extract_number_before_hyphen(s) :
    match = re.search(r'\d[^-]*', s)
    
    return match.group(0) if match else ''

def extract_nums_chars_nums(s) :
    pattern = r'\b\d+-[a-zA-Z]+-\d+\b'
    match = re.search(pattern, s)

    return match.group(0) if match else None


def extract_text_from_dxf(dxf_path) :
    doc = ezdxf.readfile(dxf_path)
    
    texts = []
    for entity in doc.modelspace().query('TEXT') :
        texts.append(entity.dxf.text)
        
    return texts


import fitz  # PyMuPDF

class ISOPdf:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.iso_drawing = fitz.open(pdf_path)
        
    def extract_text_by_coor(self, rect, page_number=0):  # 첫 번째 페이지는 0
        page = self.iso_drawing[page_number]
        text = page.get_text("text", clip=rect)
        return text
    
    def close_pdf(self):
        self.iso_drawing.close()  # 파일을 닫는 올바른 호출


import tkinter as tk
from tkinter import Canvas, Button
import fitz  # PyMuPDF

def show_pdf_coor(pdf_path):
    # Tkinter 창 생성
    root = tk.Tk()

    # PDF 파일 열기
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)  # 첫 번째 페이지 로드

    # PDF 페이지를 이미지로 변환
    pix = page.get_pixmap()
    img = tk.PhotoImage(width=pix.width, height=pix.height, master=root)
    img.put(pix.tobytes("ppm"))

    canvas = Canvas(root, width=pix.width, height=pix.height)
    canvas.pack()

    # 이미지를 캔버스에 표시
    canvas.create_image(0, 0, image=img, anchor="nw")

    # 마우스 클릭 이벤트 처리
    def on_click(event):
        x, y = event.x, event.y
        print('Clicked at: ({}, {})'.format(x, y))

    canvas.bind('<Button-1>', on_click)  # 왼쪽 마우스 버튼 클릭 이벤트 바인드

    # 종료 버튼 추가
    quit_button = Button(root, text="Quit", command=root.destroy)
    quit_button.pack()

    # 창 실행
    root.mainloop()
