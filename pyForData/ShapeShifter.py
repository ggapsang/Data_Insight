import os
import fitz
import ezdxf
import pandas as pd
from PIL import Image, ImageSequence
from tqdm import tqdm

class ShapeShifter():
    
    def __init__(self, f_path):
        self.f_path = f_path
        
    def tif_to_pdf(self, pdf_path, color=False):
        with Image.open(self.f_path) as img:
            if color:
                images = [page.convert('RGB') for page in ImageSequence.Iterator(img)]
            else:
                images = [page.convert('1') for page in ImageSequence.Iterator(img)]  # '1' for binary (black and white) images
            
        images[0].save(pdf_path, save_all=True, append_images=images[1:], format='PDF')

### Test Code
"""
tif_folder = "D:\\pseudoDB\\(ISO dwg) 속성값 추출 작업)\\00_tif\\"
pdf_folder = "D:\\pseudoDB\\(ISO dwg) 속성값 추출 작업)\\01_dwg_to_pdf\\배치_tif_to_pdf\\" # pdf 변환 파일 저장할 위치

file = "01-6000-01-S.tif"
pdf_file = "01-6000-01-S.pdf"

f_path = os.path.join(tif_folder, file)
pdf_path = os.path.join(pdf_folder, pdf_file)

shape_shifter = ShapeShifter(f_path)
shape_shifter.tif_to_pdf(pdf_path)
"""

### Batch
"""
tif_folder = "D:\\pseudoDB\\(ISO dwg) 속성값 추출 작업)\\00_tif\\"
pdf_folder = "D:\\pseudoDB\\(ISO dwg) 속성값 추출 작업)\\01_dwg_to_pdf\\배치_tif_to_pdf\\" # pdf 변환 파일 저장할 위치

tif_files = os.listdir(tif_folder)
e_f_list = []

for tif_f in tqdm(tif_files) :
    pdf_f = tif_f.replace(".tif", ".pdf")
    tif_path = os.path.join(tif_folder, tif_f)
    pdf_path = os.path.join(pdf_folder, pdf_f)
    
    shape_shifter = ShapeShifter(tif_path)
    try :
        shape_shifter.tif_to_pdf(pdf_path, color=False)
    except :
        e_f_list.append(tif_f)
"""
