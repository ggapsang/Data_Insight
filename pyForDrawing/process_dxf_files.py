import re
import os
import ezdxf

def extract_text_from_dxf(dxf_path) :
    """dxf 파일에서 텍스트 추출"""
    doc = ezdxf.readfile(dxf_path)
    
    texts = []
    for entity in doc.modelspace().query('TEXT') :
        texts.append(entity.dxf.text)
        
    return texts