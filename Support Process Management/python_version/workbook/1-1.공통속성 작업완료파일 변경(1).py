# %%
import numpy as np
import pandas as pd
import os
import sys
from datetime import datetime
from tqdm import tqdm

import PivotTable as pt
from process_mangement import TableTransformer as tt

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

# 분류체계 파일 로드
f_cct = "분류체계.csv"
path_cct = os.path.join(folder_schema, f_cct)
df_cct = pd.read_csv(path_cct)

# 속성체계 파일 로드
f_indiv = "속성체계.csv"
path_indiv = os.path.join(folder_schema, f_indiv)
df_indiv = pd.read_csv(path_indiv)

# 공통속성 작업파일 로드
f_nm = "2201_521,524,525,527,529_공통작업파일.csv"
f_path = os.path.join(folder_common, f_nm)
df_common = pd.read_csv(f_path, encoding='utf8')

# 해더 파일 로드
header_path = os.path.join(folder_schema, "info_headers.csv")
df_headers = pd.read_csv(header_path, encoding='utf8')
df_headers = df_headers.fillna(value=np.nan)

# 표준데이터시트 cct 파일 로드
f_std_ds = "표준데이터시트(cct).csv"
f_std_ds_path = os.path.join(folder_src, f_std_ds)
df_std_ds_cct = pd.read_csv(f_std_ds_path, encoding='utf8')

# 비표준데이터시트 cct 파일 로드
f_nonstd_ds = "비표준데이터시트(cct).csv"
f_nonstd_ds_path = os.path.join(folder_src, f_nonstd_ds)
df_nonstd_ds_cct = pd.read_csv(f_nonstd_ds_path, encoding='utf8')

# %%
transformer = tt(df_common, df_cct, df_indiv)

# %%
df_indiv = transformer.from_common_to_indiv(df_headers=df_headers, df_nonstd_ds_cct=df_nonstd_ds_cct, df_std_ds_cct=df_std_ds_cct)
df_indiv.to_csv(os.path.join(folder_result, f"(개별속성)_{process}_추가분-521,524,525,527,529_데이터 정비 파이프라인_개별속성 작업 시작_{today}.csv"), index=False, encoding='cp949')


