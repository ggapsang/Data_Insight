import ezdxf
import pandas as pd
import os
import shutil
from tqdm import tqdm
from concurrent.futures import ProcessPoolExecutor
import numpy as np
import warnings
import json
warnings.filterwarnings('ignore')


############ ISO Crawling 기본 함수 ############

def extract_text_from_entities(entities, doc, data):
    """엔티티 목록에서 텍스트 추출"""
    for entity in entities:
        if entity.dxftype() in ['TEXT', 'MTEXT', 'ATTDEF', 'ATTRIB']:
            data.append({
                "Type": entity.dxftype(),
                "Text": entity.dxf.text if hasattr(entity.dxf, 'text') else entity.text,
                "Position_X": entity.dxf.insert.x,
                "Position_Y": entity.dxf.insert.y
            })
        elif entity.dxftype() == 'INSERT':
            # 중첩된 블록 참조 처리
            nested_block = doc.blocks.get(entity.dxf.name)
            extract_text_from_entities(nested_block, doc, data)

def extract_all_text_in_dxf_advance(dxf_file):
    """DXF 파일의 모든 레이아웃과 레이어에서 텍스트 엔티티 추출"""
    try:
        doc = ezdxf.readfile(dxf_file)
    except IOError as e:
        print(f"File cannot be opened: {e}")
        return None
    except ezdxf.DXFStructureError as e:
        print(f"DXF structure error: {e}")
        return None

    data = []
    # 모든 레이아웃 탐색 (모델 공간 및 페이퍼 공간)
    for layout in doc.layouts:
        extract_text_from_entities(layout, doc, data)

    df = pd.DataFrame(data)
    return df

def find_next_text_in_x_direction_limit_length(df, reference_text, tolerance=3, max_distance=10):
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

def find_next_text_in_y_direction_limit_length(df, reference_text, tolerance=3, max_distance=10):
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




############ ISO Crawling Main 함수 ############

def process_extract_text_on_keyword_1(dxf_f_list, path, tolerance=3):
    col_f_nm = []
    col_pwht = []
    col_ndt = []
    col_test_press = []
    col_operating_press = []
    col_operating_temp = []
    col_design_press = []
    col_design_temp = []
    col_thk = []

    for dxf_f in tqdm(dxf_f_list):
        col_f_nm.append(dxf_f)  # 파일명을 리스트에 추가

        dxf_file = os.path.join(path, dxf_f)
        df = extract_all_text_in_dxf_advance(dxf_file)

        # PWHT
        pwht = find_next_text_in_x_direction(df, "PWHT", tolerance)
        col_pwht.append(pwht if pwht is not None else None)

        # NDT
        ndt = find_next_text_in_x_direction(df, "NDT", tolerance)
        col_ndt.append(ndt if ndt is not None else None)

        # TEST PRESSURE
        test_press = find_next_text_in_x_direction(df, "TEST", 3)
        if test_press is None :
            test_press = find_next_text_in_x_direction(df, "HYDRO", 4)
        col_test_press.append(test_press if test_press is not None else None)

        # OPERATING PRESSURE
        operating_press = find_next_text_in_x_direction(df, "OPERATING", tolerance)
        col_operating_press.append(operating_press if operating_press is not None else None)

        # DESIGN PRESSURE
        design_press = find_next_text_in_x_direction(df, "DESIGN", tolerance)
        col_design_press.append(design_press if design_press is not None else None)

        # DESIGN TEMPERATRUE
        design_temp = find_next_text_in_y_direction(df, "TEMP", tolerance)
        col_design_temp.append(design_temp if design_temp is not None else None)

        # OPERATING TEMPERATURE
        operating_temp = find_next_text_in_y_direction(df, design_temp, tolerance)
        col_operating_temp.append(operating_temp if operating_temp is not None else None)

        # INSULATION THICKNESS
        thk = find_next_text_in_y_direction(df, "THK", 4)
        col_thk.append(thk if thk is not None else None)

    # 데이터 프레임 생성 전 길이 확인
    print("Length Check:")
    print("File Names:", len(col_f_nm))
    print("PWHT:", len(col_pwht))
    print("NDT:", len(col_ndt))
    print("Test Pressure:", len(col_test_press))
    print("Operating Pressure:", len(col_operating_press))
    print("Operating Temperature:", len(col_operating_temp))
    print("Design Pressure:", len(col_design_press))
    print("Design Temperature:", len(col_design_temp))
    print("Thickness:", len(col_thk))

    result_df = pd.DataFrame({
        "파일": col_f_nm,
        "PWHT": col_pwht,
        "NDT": col_ndt,
        "TEST": col_test_press,
        "OPERATING PRESS": col_operating_press,
        "OPERATING TEMP": col_operating_temp,
        "DESIGN PRESS": col_design_press,
        "DESIGN TEMP": col_design_temp,
        "THK": col_thk
    })

    return result_df

def process_extract_text_on_keyword_2(dxf_f_list, path, tolerance=3):
    col_f_nm = []
    col_pwht = []
    col_ndt = []
    col_test_press = []
    col_operating_press = []
    col_operating_temp = []
    col_design_press = []
    col_design_temp = []
    col_thk = []

    for dxf_f in tqdm(dxf_f_list):
        col_f_nm.append(dxf_f)  # 파일명을 리스트에 추가

        dxf_file = os.path.join(path, dxf_f)
        df = extract_all_text_in_dxf_advance(dxf_file)

        # PWHT
        pwht = find_next_text_in_x_direction(df, "PWHT", tolerance)
        col_pwht.append(pwht if pwht is not None else None)

        # NDT
        ndt = find_next_text_in_x_direction(df, "NDT", tolerance)
        col_ndt.append(ndt if ndt is not None else None)

        # TEST PRESSURE
        test_press = find_next_text_in_x_direction(df, "TEST", 3)
        if test_press is None :
            test_press = find_next_text_in_x_direction(df, "HYDRO", 4)
        col_test_press.append(test_press if test_press is not None else None)

        # OPERATING PRESSURE
        operating_press = find_next_text_in_x_direction(df, "OPERATING", tolerance)
        col_operating_press.append(operating_press if operating_press is not None else None)

        # DESIGN PRESSURE
        design_press = find_next_text_in_x_direction(df, "DESIGN", tolerance)
        col_design_press.append(design_press if design_press is not None else None)

        # DESIGN TEMPERATRUE
        design_temp = find_next_text_in_x_direction(df, design_press, tolerance)
        col_design_temp.append(design_temp if design_temp is not None else None)

        # OPERATING TEMPERATURE
        operating_temp = find_next_text_in_y_direction(df, design_temp, tolerance)
        col_operating_temp.append(operating_temp if operating_temp is not None else None)

        # INSULATION THICKNESS
        thk = find_next_text_in_y_direction(df, "THK", 4)
        col_thk.append(thk if thk is not None else None)
  
    # 데이터 프레임 생성 전 길이 확인
    print("Length Check:")
    print("File Names:", len(col_f_nm))
    print("PWHT:", len(col_pwht))
    print("NDT:", len(col_ndt))
    print("Test Pressure:", len(col_test_press))
    print("Operating Pressure:", len(col_operating_press))
    print("Operating Temperature:", len(col_operating_temp))
    print("Design Pressure:", len(col_design_press))
    print("Design Temperature:", len(col_design_temp))
    print("Thickness:", len(col_thk))

    result_df = pd.DataFrame({
        "파일": col_f_nm,
        "PWHT": col_pwht,
        "NDT": col_ndt,
        "TEST": col_test_press,
        "OPERATING PRESS": col_operating_press,
        "OPERATING TEMP": col_operating_temp,
        "DESIGN PRESS": col_design_press,
        "DESIGN TEMP": col_design_temp,
        "THK": col_thk
    })

    return result_df


def crawling_by_json_coor(dxf_f_list, path, json_coor_path, tolerance=3) :
    """
    동일한 템플릿에서 추출할 텍스트의 좌표계를 미리 저장해 둔 json 파일을 읽어들여 해당 좌표의 텍스트를 추출
    param :
        - dxf_f_list (list) : 파싱할 .dxf 파일들의 리스트
        - path (str) : 파싱할 dxf 파일들이 있는 폴더 경로
        - json_coor_path (str) : json 파일. 경로와 파일명을 모두 포함함
        - tolerance : 각 좌표 지점에 정확히 맞지 않더라도 추출할 오차 범위를 지정함 default 값은 +-3
    return :
        - df (DataFrame) : 파일별로 속성값이 추출되어 있는 데이터프레임을 반환한다  
    """

    col_f_nm = []
    col_pwht = []
    col_ndt = []
    col_test_press = []
    col_operating_press = []
    col_operating_temp = []
    col_design_press = []
    col_design_temp = []
    col_thk = []

    # json 파일 읽어들이기

    with open(json_coor_path, 'r') as file :
        coor = json.load(file)

    pwht_x = coor["PWHT"]["X"]
    pwht_y = coor["PWHT"]["Y"]
    ndt_x = coor["NDT"]["X"]
    ndt_y = coor["NDT"]["Y"]
    test_press_x = coor["TEST PRESSURE"]["X"]
    test_press_y = coor["TEST PRESSURE"]["Y"]
    operating_press_x = coor["OPERATING PRESSURE"]["X"]
    operating_press_y = coor["OPERATING PRESSURE"]["Y"]
    design_press_x =  coor["DESIGN PRESSURE"]["X"]
    design_press_y = coor["DESIGN PRESSURE"]["Y"]
    design_temp_x = coor["DESIGN TEMPERATURE"]["X"]
    design_temp_y = coor["DESIGN TEMPERATURE"]["Y"]
    operating_temp_x = coor["OPERATING TEMPERATURE"]["X"]
    operating_temp_y = coor["OPERATING TEMPERATURE"]["Y"]
    thk_x = coor["INSULATION THICKNESS"]["X"]
    thk_y = coor["INSULATION THICKNESS"]["Y"]

    for dxf_f in tqdm(dxf_f_list):
        col_f_nm.append(dxf_f)  # 파일명을 리스트에 추가

        dxf_file = os.path.join(path, dxf_f)
        df = extract_all_text_in_dxf_advance(dxf_file)

        # PWHT
        pwht = find_closest_text(df, pwht_x, pwht_y, tolerance)
        col_pwht.append(pwht if pwht is not None else None)

        # NDT
        ndt = find_closest_text(df, ndt_x, ndt_y, tolerance)
        col_ndt.append(ndt if ndt is not None else None)

        # TEST PRESSURE
        test_press = find_closest_text(df, test_press_x, test_press_y, tolerance)
        col_test_press.append(test_press if test_press is not None else None)

        # OPERATING PRESSURE
        operating_press = find_closest_text(df, operating_press_x, operating_press_y, tolerance)
        col_operating_press.append(operating_press if operating_press is not None else None)

        # DESIGN PRESSURE
        design_press = find_closest_text(df, design_press_x, design_press_y, tolerance)
        col_design_press.append(design_press if design_press is not None else None)

        # DESIGN TEMPERATRUE
        design_temp = find_closest_text(df, design_temp_x, design_temp_y)
        col_design_temp.append(design_temp if design_temp is not None else None)

        # OPERATING TEMPERATURE
        operating_temp = find_closest_text(df, operating_temp_x, operating_temp_y)
        col_operating_temp.append(operating_temp if operating_temp is not None else None)
        
        # INSULATION THICKNESS
        thk = find_closest_text(df, thk_x, thk_y)
        col_thk.append(thk if thk is not None else None)

    # 데이터 프레임 생성 전 길이 확인
    print("Length Check:")
    print("File Names:", len(col_f_nm))
    print("PWHT:", len(col_pwht))
    print("NDT:", len(col_ndt))
    print("Test Pressure:", len(col_test_press))
    print("Operating Pressure:", len(col_operating_press))
    print("Operating Temperature:", len(col_operating_temp))
    print("Design Pressure:", len(col_design_press))
    print("Design Temperature:", len(col_design_temp))
    print("Thickness:", len(col_thk))

    result_df = pd.DataFrame({
        "파일": col_f_nm,
        "PWHT": col_pwht,
        "NDT": col_ndt,
        "TEST": col_test_press,
        "OPERATING PRESS": col_operating_press,
        "OPERATING TEMP": col_operating_temp,
        "DESIGN PRESS": col_design_press,
        "DESIGN TEMP": col_design_temp,
        "THK": col_thk
    })        





############ 실제 처리(사용법) ############

dxf_folder_path = f"D:\\(Supporter)\\dxf 태그 추출\\{process_no}_dxf\\"
dxf_f_list = os.listdir(dxf_folder_path)
path = dxf_folder_path

json_coor_path = "D:\\(Supporter)\\dxf 태그 추출\\템플릿1_좌표.json"
result_df_coor = process_extract_text_on_coor(dxf_f_list, path, json_coor_path)
result_df_key1 = process_extract_text_on_keyword_1(dxf_f_list, path)
result_df_key2 = process_extract_text_on_keyword_2(dxf_f_list, path)

result_df_key1.to_csv(f"{process_no}_추출_key1.csv", encoding='utf-8-sig', index=False)
result_df_key2.to_csv(f"{process_no}_추출_key2.csv", encoding='utf-8-sig', index=False)
result_df_coor.to_csv(f"{process_no}_추출_좌표.csv", encoding='utf-8-sig', index=False)
