import re
import os
import ezdxf

def extract_text_from_dxf(dxf_path) :
    doc = ezdxf.readfile(dxf_path)
    
    texts = []
    for entity in doc.modelspace().query('TEXT') :
        texts.append(entity.dxf.text)
        
    return texts
