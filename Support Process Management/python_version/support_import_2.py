# %%
import os
import pandas as pd
import numpy as np
import sqlite3
import json

from tqdm import tqdm
from datetime import datetime

tqdm.pandas()

import warnings
warnings.filterwarnings("ignore")

# %%
def create_nested_dataframe(df, filter_col, filter_val, header_col, header_val):
    """
    주어진 데이터프레임에서 특정 필터 조건에 맞는 데이터를 먼저 추출하고,
    다시 해당 조건 중 값을 가진 레코드를 헤더로 사용하는 중첩 데이터프레임을 생성

    Parameters:
    df (pd.DataFrame): 원본 데이터프레임
    filter_col (str): 필터링에 사용할 컬럼명
    filter_val (any): 필터링 조건에 사용할 값
    header_col (str): 헤더로 사용할 레코드를 찾기 위한 컬럼명
    header_val (any): 헤더로 사용할 레코드를 찾기 위한 값

    Returns:
    pd.DataFrame: 헤더를 설정한 새로운 중첩 데이터프레임
    """


    # 필터링 조건에 맞는 데이터프레임 생성
    filtered_df = df[df[filter_col] == filter_val]

    # 헤더로 사용할 레코드 찾기
    header = filtered_df[filtered_df[header_col] == header_val].iloc[0]

    # 나머지 데이터를 새로운 데이터프레임으로 구성
    nested_df = filtered_df[filtered_df[header_col] != header_val].reset_index(drop=True)

    # 새 데이터프레임의 컬럼명을 헤더로 설정
    nested_df.columns = header

    # 'Unnamed'로 시작하는 모든 열을 드롭
    nested_df = nested_df.loc[:, ~nested_df.columns.astype(str).str.startswith('Unnamed')]
    nested_df = nested_df.loc[:, ~nested_df.columns.astype(str).str.contains('nan', case=False)]

    return nested_df

# %%
## step 0 : 경로 저장

db_path = os.path.join(os.getcwd(), '개별속성.sqlite3')
output_path = os.path.join(os.getcwd(), "2202_import_test.xlsx")

# %%
## step 1 : 데이터베이스에 접속하여 테이블을 읽어 데이터프레임으로 가져온다
conn = sqlite3.connect(db_path)
indiv_table_name = "2201_개별속성리스트"
common_table_name  = '2201_공통속성'
df_indiv = pd.read_sql_query(f'SELECT * FROM "{indiv_table_name}"', conn)
df_common = pd.read_sql_query(f'SELECT * FROM "{common_table_name}"', conn)

# %%
## step2 : 사전 저장된 산출물 템플릿 파일을 불러온다
df_indiv_templates = pd.read_csv('개별속성템플릿_v2.51.csv')
with open ("cct_dict_240612.json") as f :
    cct_json = json.load(f)

# %%
def load_renaming_mapping(cct_json, cct) :
    """각 속성 항목이 개별속성 몇 번째에 매핑되는지를 리스트로 반환"""
    cct_idx = cct_json["index"][cct]
    cct_dict = cct_json['header_list'][cct_idx]
    try :
        del cct_dict["C|C|T"]
    except :
        pass
    key_attr_no = cct_dict.keys()
    val_attr_nm = cct_dict.values()
    reversed_nm_dict = dict(zip(val_attr_nm, key_attr_no))

    return reversed_nm_dict

# %%
cct = "FIXED EQUIPMENT|VESSEL|HORIZONTAL"
nm_mapping = load_renaming_mapping(cct_json, cct)

# %%
df_indiv_template = df_indiv_templates[df_indiv_templates['C|C|T'] =="FIXED EQUIPMENT|VESSEL|HORIZONTAL"]
df_indiv_template = df_indiv_template[df_indiv_template['속성 그룹 코드']=='01_속성명']

# %%
df_pivot_ready = df_indiv[df_indiv['C|C|T'] == "FIXED EQUIPMENT|VESSEL|HORIZONTAL"]
df_pivot = df_pivot_ready[['SRNo', '속성명', '속성값']].pivot(index='SRNo', columns='속성명', values='속성값')
df_pivot_ren = df_pivot.rename(columns=nm_mapping)
df_pivot_ren

# %%
df_indiv_template

# %%
df = pd.concat([df_indiv_template, df_pivot_ren], axis=0)
df

# %%
df.iloc[:, 20:30]

# %%



