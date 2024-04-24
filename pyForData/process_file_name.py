import os

def get_file_name_strip(file) :
    """확장자를 제외한 파일명 추출"""
    basename = os.path.basename(file)
    file_name = os.path.splitext(basename)[0]
    
    return file_name

def get_file_extension(file) :
    """확장자 추출"""
    basename = os.path.basename(file)
    file_name = os.path.splitext(basename)[1]
    
    return file_name
