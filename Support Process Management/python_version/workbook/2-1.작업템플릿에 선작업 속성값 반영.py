# %%
import pandas as pd
import numpy as np
import os

import PivotTable as pt
import process_mangement as pm

import warnings
warnings.filterwarnings("ignore")

# %%
def make_header_to_firts_row(df) :
    """해더를 첫번째 행으로 이동"""
    header_df = pd.DataFrame([df.columns.tolist()], columns=df.columns)
    new_df = pd.concat([header_df, df], ignore_index=True)
    return new_df

def cut_header(df, df_cut) :
    """해더 길이 맞추기"""
    main_header = df.columns.to_list()
    main_header = main_header[: len(df_cut.columns.to_list())]
    return main_header

def change_header(df, main_header) :
    """메인 해더로 해더 이름 바꾸기"""
    df.columns = main_header
    return df

# %% [markdown]
# ##### 표준데이터시트

# %%
folder_base = os.getcwd()
folder_common = os.path.join(folder_base, "working_common")
folder_schema = os.path.join(folder_base, "schema")
folder_src = os.path.join(folder_base, "working_by_source")
folder_result = os.path.join(folder_base, "result")

db_path = os.path.join(folder_base, "GSC_INDIV.sqlite")

readDb = pm.ReadDB(db_path)

df_attrs_stdds = readDb.read_db_to_dataframe('개별속성_표준데이터시트')
df_attrs_stdds['공종별분류코드'].astype(float).astype(int)
# df_attrs_nonstdds = readDb.read_db_to_dataframe('비표준데이터시트_개별속성')

df_attrs_header = pd.read_csv(os.path.join(folder_schema, '속성해더.csv'), dtype={'공정번호' : str}) # 속성해더
df_attrs_schema = pd.read_csv(os.path.join(folder_schema, '속성체계(속성순번).csv'))

# %%
working_f_nm = "(개별속성)_2201_추가분-520_데이터 정비 파이프라인_개별속성 작업 시작_20240507.csv"
df_working_f = pd.read_csv(os.path.join(folder_base, working_f_nm), encoding='utf8')

# %%
df_attrs_header = df_attrs_header[df_attrs_header['속성 그룹 코드']=='01_속성명']
insert_attrs_preprocessing = pm.InsertAttrstPreprocessing(df_working_f, df_attrs_header)
insert_attrs_preprocessing.step0_1()
source = '표준데이터시트'
insert_attrs_preprocessing.step0_2(df_attrs_stdds, col_name=source)
insert_attrs_preprocessing.setp0_3(df_attrs_header)

df_working_f = insert_attrs_preprocessing.df_working_f
df_attrs = insert_attrs_preprocessing.df_attrs
df_attrs_header = insert_attrs_preprocessing.df_attrs_headers

# %%
df_attrs_schema = df_attrs_schema[['공종별분류코드', '속성명', '속성']]
df_attrs_schema['공종별분류코드'].astype(float).astype(int)
df_attrs_add = pd.merge(df_attrs, df_attrs_schema, left_on=['공종별분류코드', '속성명'], right_on=['공종별분류코드', '속성명'])
df_attrs_add.head()

# %%
df_attrs_add_filtered = df_attrs_add[['SRNo', '속성값', '속성']]
df_attrs_add_filtered.head()

# %%
df_attrs_add_filtered.drop_duplicates(subset=['SRNo', '속성'], inplace=True)

# %%
pivot_df = df_attrs_add_filtered.pivot(index='SRNo', columns='속성', values='속성값')
pivot_df.reset_index(inplace=True)

pivot_df.rename(columns={'SRNo' : '표준데이터시트'}, inplace=True)

# %%
header_sorted = df_working_f.columns.to_list()
merge_df = pd.merge(df_working_f, pivot_df, how='left', on='표준데이터시트', suffixes=['_drop', ''])
merge_df = merge_df[[col for col in merge_df.columns.to_list() if '_drop' not in col]]
result_df = merge_df[header_sorted]

# %%
result_df.to_csv('check.csv', encoding='cp949', index=False)

# %% [markdown]
# ##### 비표준데이터시트

# %%
process_no = 2102
folder_base = os.getcwd()
folder_common = os.path.join(folder_base, "working_common")
folder_schema = os.path.join(folder_base, "schema")
folder_src = os.path.join(folder_base, "working_by_source")
folder_result = os.path.join(folder_base, "result")

db_path = os.path.join(folder_base, "GSC_INDIV.sqlite")

readDb = pm.ReadDB(db_path)

df_attrs_stdds = readDb.read_db_to_dataframe(f'개별속성_{process_no}')
df_attrs_stdds['공종별분류코드'].astype(float).astype(int)
# df_attrs_nonstdds = readDb.read_db_to_dataframe('비표준데이터시트_개별속성')

df_attrs_header = pd.read_csv(os.path.join(folder_schema, '속성해더.csv'), dtype={'공정번호' : str}) # 속성해더
df_attrs_schema = pd.read_csv(os.path.join(folder_schema, '속성체계(속성순번).csv'))

# %%
working_f_nm = "(개별속성)_2201_추가분-520_데이터 정비 파이프라인_개별속성 작업 시작_20240507.csv"
df_working_f = pd.read_csv(os.path.join(folder_base, working_f_nm), encoding='utf8')

# %%
df_attrs_header = df_attrs_header[df_attrs_header['속성 그룹 코드']=='01_속성명']
insert_attrs_preprocessing = pm.InsertAttrstPreprocessing(df_working_f, df_attrs_header)
insert_attrs_preprocessing.step0_1()
source = '선작업 태그'
insert_attrs_preprocessing.step0_2(df_attrs_stdds, col_name=source)
insert_attrs_preprocessing.setp0_3(df_attrs_header)

df_working_f = insert_attrs_preprocessing.df_working_f
df_attrs = insert_attrs_preprocessing.df_attrs
df_attrs_header = insert_attrs_preprocessing.df_attrs_headers

# %%
df_attrs_schema = df_attrs_schema[['공종별분류코드', '속성명', '속성']]
df_attrs_schema['공종별분류코드'].astype(float).astype(int)
df_attrs_add = pd.merge(df_attrs, df_attrs_schema, left_on=['공종별분류코드', '속성명'], right_on=['공종별분류코드', '속성명'])
df_attrs_add.head()

# %%
df_attrs_add_filtered.drop_duplicates(subset=['SRNo', '속성'], inplace=True)

# %%
pivot_df = df_attrs_add_filtered.pivot(index='SRNo', columns='속성', values='속성값')
pivot_df.reset_index(inplace=True)

pivot_df.rename(columns={'SRNo' : '표준데이터시트'}, inplace=True)

# %%
header_sorted = df_working_f.columns.to_list()
merge_df = pd.merge(df_working_f, pivot_df, how='left', on='표준데이터시트', suffixes=['_drop', ''])
merge_df = merge_df[[col for col in merge_df.columns.to_list() if '_drop' not in col]]
result_df = merge_df[header_sorted]

# %%
result_df.to_csv('check.csv', encoding='cp949', index=False)


