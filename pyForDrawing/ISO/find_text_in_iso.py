import ezdxf
import pandas as pd
import os
import shutil
from tqdm import tqdm
from concurrent.futures import ProcessPoolExecutor

def extract_all_text_in_dxf(dxf_file):
    doc = ezdxf.readfile(dxf_file)
    data = []
    
    # 모델 공간에서 모든 텍스트 관련 엔티티를 반복 처리
    msp = doc.modelspace()
    for entity in msp:
        if entity.dxftype() == 'TEXT':
            data.append({
                "Type": entity.dxftype(),
                "Text": entity.dxf.text,
                "Position_X": entity.dxf.insert.x,
                "Position_Y": entity.dxf.insert.y
            })
        elif entity.dxftype() == 'MTEXT':
            data.append({
                "Type": entity.dxftype(),
                "Text": entity.text,
                "Position_X": entity.dxf.insert.x,
                "Position_Y": entity.dxf.insert.y
            })
        elif entity.dxftype() in ['ATTDEF', 'ATTRIB']:
            data.append({
                "Type": entity.dxftype(),
                "Text": entity.dxf.text,
                "Position_X": entity.dxf.insert.x,
                "Position_Y": entity.dxf.insert.y
            })
        elif entity.dxftype() == 'INSERT':
            # INSERT 엔티티는 블록 참조로, 블록 내 텍스트 엔티티를 추출
            block = doc.blocks.get(entity.dxf.name)
            for block_entity in block:
                if block_entity.dxftype() in ['TEXT', 'MTEXT', 'ATTDEF', 'ATTRIB']:
                    data.append({
                        "Type": block_entity.dxftype(),
                        "Text": block_entity.dxf.text if block_entity.dxftype() != 'MTEXT' else block_entity.text,
                        "Position_X": block_entity.dxf.insert.x,
                        "Position_Y": block_entity.dxf.insert.y
                    })
    
    # 데이터를 DataFrame으로 변환
    df = pd.DataFrame(data)
    
    return df

# 특정 텍스트를 기준으로 X 방향으로 가장 먼저 나오는 텍스트 찾기
def find_next_text_in_x_direction(df, reference_text, tolerance=3):
    """특정 텍스트를 기준으로 x 방향으로 가장 먼저 나오는 텍스트 찾기"""
    # 기준 텍스트의 위치 찾기
    ref_row = df[df['Text'] == reference_text]
    if ref_row.empty:
        return None
    
    ref_x = ref_row.iloc[0]['Position_X']
    ref_y = ref_row.iloc[0]['Position_Y']
    
    # 기준 텍스트의 X 좌표보다 큰 X 좌표를 가진 텍스트 중 가장 가까운 텍스트 찾기 (오차 범위 고려)
    candidates = df[(df['Position_X'] > ref_x) & (df['Position_Y'].between(ref_y - tolerance, ref_y + tolerance))]
    if candidates.empty:
        return None
    
    next_text_row = candidates.loc[candidates['Position_X'].idxmin()]
    return next_text_row['Text']

# 특정 텍스트를 기준으로 Y 방향으로 가장 먼저 나오는 텍스트 찾기
def find_next_text_in_y_direction(df, reference_text, tolerance=3):
    """특정 텍스트를 기준으로  y 방향으로 가장 먼저 나오는 텍스트 찾기"""
    # 기준 텍스트의 위치 찾기
    ref_row = df[df['Text'] == reference_text]
    if ref_row.empty:
        return None
    
    ref_x = ref_row.iloc[0]['Position_X']
    ref_y = ref_row.iloc[0]['Position_Y']
    
    # 기준 텍스트의 Y 좌표보다 큰 Y 좌표를 가진 텍스트 중 가장 가까운 텍스트 찾기 (오차 범위 고려)
    candidates = df[(df['Position_Y'] > ref_y) & (df['Position_X'].between(ref_x - tolerance, ref_x + tolerance))]
    if candidates.empty:
        return None
    
    next_text_row = candidates.loc[candidates['Position_Y'].idxmin()]
    return next_text_row['Text']

def find_next_text_in_x_direction_advance(df, reference_text, tolerance=3, max_distance=10):
    """특정 텍스트를 기준으로 x 방향으로 가장 먼저 나오는 텍스트를 찾되, 최대 거리 내에서만 찾기"""
    ref_row = df[df['Text'] == reference_text]
    if ref_row.empty:
        return None
    
    ref_x = ref_row.iloc[0]['Position_X']
    ref_y = ref_row.iloc[0]['Position_Y']
    
    candidates = df[(df['Position_X'] > ref_x) & (df['Position_X'] <= ref_x + max_distance) & 
                    (df['Position_Y'].between(ref_y - tolerance, ref_y + tolerance))]
    if candidates.empty:
        return None
    
    next_text_row = candidates.loc[candidates['Position_X'].idxmin()]
    return next_text_row['Text']

def find_next_text_in_y_direction_advance(df, reference_text, tolerance=3, max_distance=10):
    """특정 텍스트를 기준으로 y 방향으로 가장 먼저 나오는 텍스트를 찾되, 최대 거리 내에서만 찾기"""
    ref_row = df[df['Text'] == reference_text]
    if ref_row.empty:
        return None
    
    ref_x = ref_row.iloc[0]['Position_X']
    ref_y = ref_row.iloc[0]['Position_Y']
    
    candidates = df[(df['Position_Y'] > ref_y) & (df['Position_Y'] <= ref_y + max_distance) & 
                    (df['Position_X'].between(ref_x - tolerance, ref_x + tolerance))]
    if candidates.empty:
        return None
    
    next_text_row = candidates.loc[candidates['Position_Y'].idxmin()]
    return next_text_row['Text']


def process_extract_text(dxf_f_list, tolerance=2) :

    col_f_nm = []
    col_pwht = []
    col_ndt = []
    col_test_press = []
    col_operating_press = []
    col_operating_temp = []
    col_design_press = []
    col_design_temp = []
    col_thk = []

    for dxf_f in tqdm(dxf_f_list) :
        col_f_nm.append(dxf_f.replace(".pdf", ""))

        dxf_file = os.path.join(dest_path, dxf_f)

        df = extract_all_text_in_dxf(dxf_file)

        pwht = find_next_text_in_x_direction(df, "PWHT", tolerance=tolerance)
        ndt = find_next_text_in_x_direction(df, "NDT", tolerance=tolerance)
        test_press = find_next_text_in_x_direction(df, "TEST", tolerance=tolerance)
        operating_press = find_next_text_in_x_direction(df, "OPERATING", tolerance=tolerance)
        design_press = find_next_text_in_x_direction(df, "DESIGN", tolerance=tolerance)
        
        thk = find_next_text_in_y_direction(df, "THK", tolerance=tolerance)

        col_pwht.append(pwht)
        col_ndt.append(ndt)
        col_test_press.append(test_press)

        

        col_operating_press.append(operating_press)
        col_design_press.append(design_press)
        col_thk.append(thk)

        # operating_press 값으로 operating_temp를 찾기
        if col_operating_press and col_operating_press[-1] is not None:
            operating_temp = find_next_text_in_x_direction(df, col_operating_press[-1], tolerance=tolerance)
        else:
            operating_temp = None  

        # design_press 값으로 design_temp를 찾기

        if design_press and col_design_press[-1] is not None:
            design_temp = find_next_text_in_x_direction(df, col_design_press[-1], tolerance=tolerance)
        else:
            design_temp = None  

        col_operating_temp.append(operating_temp)
        col_design_temp.append(design_temp)

    result_df = pd.DataFrame({"파일" : col_f_nm,
                              "PWHT" : col_pwht,
                              "NDT" : col_ndt,
                              "TEST" : col_test_press,
                              "OPERATING PRESS" : col_operating_press,
                              "OPERATING TEMP" : col_operating_temp,
                              "DESIGN PRESS" : col_design_press,
                              "DESIGN TEMP" : col_design_temp,
                              "THK" : col_thk})
    
    return result_df



## dxf 파일에서 텍스트 추출
dxf_folder_path = "D:\\(Supporter)\\dxf 태그 추출\\2104_dxf\\"
dxf_f_list = os.listdir(dxf_folder_path)

dxf_file = "553-NG-0035.dxf"
dxf_path = os.path.join(dest_path, dxf_file)
df = extract_all_text_in_dxf(dxf_path)
print(df.head())

## 파일 리스트가 있는 경우
rextrcat_list = pd.read_csv("2104_재추출_리스트_4차.csv")
rextrcat_list = rextrcat_list["파일"].to_list()
print(rextrcat_list)


result_df = process_extract_text(rextrcat_list)
result_df.to_csv("테스트 5차_toleracne=4.csv", encoding='utf-8-sig', index=False)


print(reuslt_df)
