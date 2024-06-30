import os
import pandas as pd
import ezdxf
from tqdm import tqdm

class IsoCrawler:
    def __init__(self, dxf_file):
        self.dxf_file = dxf_file
        self.doc = ezdxf.readfile(dxf_file)
        self.data = []

    def extract_all_text_in_dxf(self):
        """모델 공간에서 모든 텍스트 관련 엔티티를 반복 처리"""
        msp = self.doc.modelspace()
        for entity in msp:
            if entity.dxftype() == 'TEXT':
                self.data.append({
                    "Type": entity.dxftype(),
                    "Text": entity.dxf.text,
                    "Position_X": entity.dxf.insert.x,
                    "Position_Y": entity.dxf.insert.y
                })

            elif entity.dxftype() == 'MTEXT':
                self.data.append({
                    "Type": entity.dxftype(),
                    "Text": entity.text,
                    "Position_X": entity.dxf.insert.x,
                    "Position_Y": entity.dxf.insert.y
                })
            elif entity.dxftype() in ['ATTDEF', 'ATTRIB']:
                self.data.append({
                    "Type": entity.dxftype(),
                    "Text": entity.dxf.text,
                    "Position_X": entity.dxf.insert.x,
                    "Position_Y": entity.dxf.insert.y
                })
            elif entity.dxftype() == 'INSERT':
                # INSERT 엔티티는 블록 참조로, 블록 내 텍스트 엔티티를 추출
                block = self.doc.blocks.get(entity.dxf.name)
                for block_entity in block:
                    if block_entity.dxftype() in ['TEXT', 'MTEXT', 'ATTDEF', 'ATTRIB']:
                        self.data.append({
                            "Type": block_entity.dxftype(),
                            "Text": block_entity.dxf.text if block_entity.dxftype() != 'MTEXT' else block_entity.text,
                            "Position_X": block_entity.dxf.insert.x,
                            "Position_Y": block_entity.dxf.insert.y
                        })
            
            # 데이터를 DataFrame으로 변환
            self.df = pd.DataFrame(self.data)

            return self.df
        
    def find_next_text_in_x_direction(self, reference_text, tolerance=3, df=None):
            """특정 텍스트를 기준으로 x 방향으로 가장 먼저 나오는 텍스트 찾기"""
            if df is None:
                df = self.df
            
            # 기준 텍스트의 위치 찾기
            ref_row = self.df[self.df['Text'] == reference_text]
            if ref_row.empty:
                return None

            ref_x = ref_row['Position_X'].values[0]
            ref_y = ref_row['Position_Y'].values[0]

            # 기준 텍스트의 x 좌표보다 큰 x 좌표를 가진 텍스트 중 가장 가까운 텍스트 찾기 (오차 범위 고려)
            candidates = df[(df['Position_X'] > ref_x) & (df['Position_Y'].between(ref_y - tolerance, ref_y + tolerance))]

            if candidates.empty:
                return None
            
            next_text_row = candidates.loc[candidates['Position_X'].idxmin()]
            return next_text_row['Text']
        
    def find_next_text_in_y_direction(self, reference_text, tolerance=3, df=None):
            """특정 텍스트를 기준으로 y 방향으로 가장 먼저 나오는 텍스트 찾기"""
            if df is None:
                df = self.df
            
            # 기준 텍스트의 위치 찾기
            ref_row = df[df['Text'] == reference_text]
            if ref_row.empty:
                return None
            
            ref_x = ref_row['Position_X'].values[0]
            ref_y = ref_row['Position_Y'].values[0]
            
            # 기준 텍스트의 y 좌표보다 큰 y 좌표를 가진 텍스트 중 가장 가까운 텍스트 찾기 (오차 범위 고려)
            candidates = df[(df['Position_Y'] > ref_y) & (df['Position_X'].between(ref_x - tolerance, ref_x + tolerance))]
            
            if candidates.empty:
                return None
            
            next_text_row = candidates.loc[candidates['Position_Y'].idxmin()]
            return next_text_row['Text']
        
    def find_next_text_in_x_direction_advanced(self, reference_text, tolerance=3, max_distance=10, df=None) :
            """특정 텍스트를 기준으로 x 방향으로 가장 먼저 나오는 텍스트를 찾되, 최대 거리 내에서만 찾기"""

            if df is None:
                df = self.df

            ref_row = df[df['Text'] == reference_text]
            if ref_row.empty:
                return None
            
            ref_x = ref_row['Position_X'].values[0]
            ref_y = ref_row['Position_Y'].values[0]

            candidates = df[(df['Position_X'] > ref_x) & (df['Position_X'] <= ref_x + max_distance) &
                            (df['Position_Y'].between(ref_y - tolerance, ref_y + tolerance))]
            if candidates.empty:
                return None
            
            next_text_row = candidates.loc[candidates['Position_X'].idxmin()]
            return next_text_row['Text']
        
    def find_next_text_in_y_direction_advanced(self, reference_text, tolerance=3, max_distance=10, df=None):
            """특정 텍스트를 기준으로 y 방향으로 가장 먼저 나오는 텍스트를 찾되, 최대 거리 내에서만 찾기"""
            if df is None:
                df = self.df
            
            ref_row = df[df['Text'] == reference_text]
            if ref_row.empty:
                return None
            
            ref_x = ref_row['Position_X'].values[0]
            ref_y = ref_row['Position_Y'].values[0]
            
            candidates = df[(df['Position_Y'] > ref_y) & (df['Position_Y'] <= ref_y + max_distance) &
                            (df['Position_X'].between(ref_x - tolerance, ref_x + tolerance))]
            if candidates.empty:
                return None
            
            next_text_row = candidates.loc[candidates['Position_Y'].idxmin()]
            return next_text_row['Text']


def process_extract_text(dxf_f_list, dest_path=os.getcwd(), tolerance=2, ) :

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

        pwht = find_next_text_in_x_direction_advanced(df, "PWHT", tolerance=tolerance)
        ndt = find_next_text_in_x_direction_advanced(df, "NDT", tolerance=tolerance)
        test_press = find_next_text_in_x_direction_advanced(df, "TEST", tolerance=tolerance)
        operating_press = find_next_text_in_x_direction_advanced(df, "OPERATING", tolerance=tolerance)
        design_press = find_next_text_in_x_direction_advanced(df, "DESIGN", tolerance=tolerance)
        
        thk = find_next_text_in_y_direction_advanced(df, "THK", tolerance=tolerance)

        col_pwht.append(pwht)
        col_ndt.append(ndt)
        col_test_press.append(test_press)

        

        col_operating_press.append(operating_press)
        col_design_press.append(design_press)
        col_thk.append(thk)

        # operating_press 값으로 operating_temp를 찾기
        if col_operating_press and col_operating_press[-1] is not None:
            operating_temp = find_next_text_in_x_direction_advanced(df, col_operating_press[-1], tolerance=tolerance)
        else:
            operating_temp = None  

        # design_press 값으로 design_temp를 찾기

        if design_press and col_design_press[-1] is not None:
            design_temp = find_next_text_in_x_direction_advanced(df, col_design_press[-1], tolerance=tolerance)
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


