### jupyter notebook에서 작성되었음

# %%
import numpy as np
import pandas as pd
import os
import sys
import sqlite3
from datetime import datetime
from tqdm import tqdm


import PivotTable as pt
from process_mangement import TableTransformer as tt
import warnings
warnings.filterwarnings(action='ignore')

# %%
process = "2201"
today= datetime.today().strftime('%Y%m%d')

# %%
# 폴더경로
folder_base = os.getcwd()
folder_common = os.path.join(folder_base, "working_common")
folder_schema = os.path.join(folder_base, "schema")
folder_src = os.path.join(folder_base, "working_by_source")
folder_result = os.path.join(folder_base, "result")
folder_upload = os.path.join(folder_base, "upload")

# 분류체계 파일 로드
f_cct = "분류체계.csv"
path_cct = os.path.join(folder_schema, f_cct)
df_cct = pd.read_csv(path_cct)

# 속성체계 파일 로드
f_indiv = "속성체계.csv"
path_indiv = os.path.join(folder_schema, f_indiv)
df_indiv = pd.read_csv(path_indiv)

# 해더 파일 로드
header_path = os.path.join(folder_schema, "info_headers.csv")
df_headers = pd.read_csv(header_path, encoding='utf8')
df_headers = df_headers.fillna(value=np.nan)

# %%
# 속성작업 파일 로드
f_attrs_sta = "표준데이터시트-고정.csv"
f_attrs_rot = "표준데이터시트-회전.csv"
f_attrs_ins = "표준데이터시트-계기.csv"
f_attrs_ele = "표준데이터시트-전기.csv"
f_attrs_ect = "표준데이터시트-기타.csv"

path_attrs_sta = os.path.join(folder_src, f_attrs_sta)
path_attrs_rot = os.path.join(folder_src, f_attrs_rot)
path_attrs_ins = os.path.join(folder_src, f_attrs_ins)
path_attrs_ele = os.path.join(folder_src, f_attrs_ele)
path_attrs_ect = os.path.join(folder_src, f_attrs_ect)

df_attrs_sta = pd.read_csv(path_attrs_sta, dtype={'공정' : str, '공종별 분류코드' : int})
df_attrs_rot = pd.read_csv(path_attrs_rot, dtype={'공정' : str, '공종별 분류코드' : int})
df_attrs_ins = pd.read_csv(path_attrs_ins, dtype={'공정' : str, '공종별 분류코드' : int})
df_attrs_ele = pd.read_csv(path_attrs_ele, dtype={'공정' : str, '공종별 분류코드' : int})
df_attrs_ect = pd.read_csv(path_attrs_ect, dtype={'공정' : str, '공종별 분류코드' : int})

# %%
dfs = [df_attrs_sta, df_attrs_rot, df_attrs_ins, df_attrs_ele, df_attrs_ect]
df_total = pd.concat(dfs)

# %%
drop_list = ["자료명",
    "작업자",
    "출처",
    "파일목록",
    "선작업 태그",
    "표준데이터시트",
    "최종 CCT 변경 유무",
    "비교",
    "비고",
    "Tag No",
    "Tag No 수정",
    "카테고리",
    "클래스",
    "타입",
    "C|C|T",]

transformer = tt(df=df_total, df_cct=df_cct, df_indiv=df_indiv)
result_df = transformer.to_upload_indiv(drop_list=drop_list)
result_df.to_csv(os.path.join(folder_upload, f"upload_indiv_sqlite_{today}.csv"), index=False, encoding='utf8')

result_df['출처'] = '표준데이터시트'

result_df.rename(columns={'SR No' : 'SRNo'}, inplace=True)
result_df.reset_index(inplace=True)
result_df.rename(columns={'index' : 'ATTR_INDEX'}, inplace=True)

result_df = result_df[result_df['속성값'] !='표준데이터시트']
result_df.head()

# %%
len(result_df)

# %%


# %%
result_df.to_csv("upload_indiv_240503.csv", index=False, encoding='utf8')


