import fitz
import re
import os
import pandas as pd
import ezdxf

def extract_annotations_from_pdf(pdf_path) :
    """pdf 파일에서 주석 추출"""
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
def find_texts(re_list, annot) :
    """정규표현식을 통해 pdf 주석에서 텍스트 추출"""    
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

  # 리스트 길이 맞추기
def pad_list_to_length(original_list, target_length):
    """리스트의 길이 맞추기"""
    
    while len(original_list) < target_length:
        original_list.append(None)
    return original_list

def extract_number_before_hyphen(s) :
    """하이픈 앞의 숫자 추출"""
    match = re.search(r'\d[^-]*', s)
    
    return match.group(0) if match else ''

def extract_nums_chars_nums(s) :
    """숫자-문자-숫자 패턴 추출"""
    pattern = r'\b\d+-[a-zA-Z]+-\d+\b'
    match = re.search(pattern, s)

    return match.group(0) if match else None