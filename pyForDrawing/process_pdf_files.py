import fitz
import re
import os

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

# 정규표현식으로 추출한 주석에서 알맞는 값 찾기
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



