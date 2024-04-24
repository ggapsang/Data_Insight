import os
import fitz
import ezdxf
import pandas as pd
from PIL import Image, ImageSequence
from tqdm import tqdm
import win32.client as win32

class ShapeShifter():
    """파일 포멧을 다양하게 변환해주는 클래스"""

    def __init__(self, f_path : str):
        self.f_path = f_path
        
    def tif_to_pdf(self, pdf_path : str, color=False):
        """transfer tif file to pdf file"""

        with Image.open(self.f_path) as img:
            if color:
                images = [page.convert('RGB') for page in ImageSequence.Iterator(img)]
            else:
                images = [page.convert('1') for page in ImageSequence.Iterator(img)]  # '1' for binary (black and white) images
            
        images[0].save(pdf_path, save_all=True, append_images=images[1:], format='PDF')

    def xlsx_to_pdf(self, pdf_path : str):
        """transfer xlsx file to pdf file"""

        excel = None
        workbook = None

        try : 
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel.Workbooks.Open(self.f_path)

            workbook.ExportAsFixedFormat(0, pdf_path)
        except Exception as e:
            print(f"Error: {e}")
        finally:
            if workbook is not None:
                workbook.Close(False)
            if excel is not None :
                excel.Quit()

    def dxf_to_pdf(self, pdf_path : str):
        """transfer dxf file to pdf file"""

        pass