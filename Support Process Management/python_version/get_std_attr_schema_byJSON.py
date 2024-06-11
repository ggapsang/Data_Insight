# %%
import os
import pandas as pd
import numpy as np
import json

from tqdm import tqdm
from datetime import datetime

tqdm.pandas()

import warnings
import sqlite3
warnings.filterwarnings("ignore")

# %%
def get_cct_dictionary(df, df_cct_attrs_count, attr_nm='개별속성1', save_json=True, save_path=os.getcwd()) :
    """속성 해더를 가지고 딕셔너리를 만듬"""

    # 속성명이 있는 행만 추출
    df_headers = df[df['속성 그룹 코드'].isin(['01_속성명'])]
    attr1_col_no = df_headers.columns.get_loc(attr_nm)
    
    df_headers = df_headers[['C|C|T'] + df_headers.columns[attr1_col_no:].to_list()]
    
    # 속성명과 순번에 대한 딕셔너리 생성
    header_list = []
    for i in tqdm(range(len(df_headers))):
        df_header = df_headers.iloc[i].dropna()
        header_list.append(df_header.to_dict())

    # 딕셔너리 생성
    idx_list = [i for i in range(len(header_list))]
    cct_list = [header_list[i]['C|C|T'] for i in range(len(header_list))]
    dict_idx = dict(zip(cct_list, idx_list))
    count_attrs = dict(zip(df_cct_attrs_count['C|C|T'], df_cct_attrs_count['속성 입력 개수']))
    
    dic = {'index' : dict_idx, 'header_list' : header_list, 'count_attrs' : count_attrs}
    # json 파일로 저장
    
    if save_json :
        import json            
        today = datetime.today().strftime("%y%m%d")
        with open(os.path.join(save_path, f'cct_dict_{today}.json'), 'w') as f:
            json.dump(dic, f)

    return dic

# %%
## step 0 : define the path of the data : json file
today = datetime.today().strftime("%Y%m%d")
work_folder_path = os.path.join(os.getcwd(), 'working_file')
result_folder_path = os.path.join(os.getcwd(), 'results')

# %%
working_f_nm = "2.5_cct_attr.csv"
work_f_path = os.path.join(work_folder_path, working_f_nm)
df_json_src = pd.read_csv(work_f_path)

# %%
dic_cct = get_cct_dictionary(df=df_json_src, df_cct_attrs_count=df_json_src)


