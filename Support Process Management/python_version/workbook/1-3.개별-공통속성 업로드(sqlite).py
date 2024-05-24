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
f_attrs = "표준데이터시트-기타.csv"
path_attrs = os.path.join(folder_src, f_attrs)
df_attrs_sta = pd.read_csv(path_attrs)
df_attrs_sta['작업자'] = None

# %%
df_headers['개별속성 공통 업로드'].replace({np.nan: None}, inplace=True)
headers = df_headers['개별속성 공통 업로드'].to_list()
headers = [col for col in headers if col != None]
transformer = tt(df=df_attrs_sta, df_cct=df_cct, df_indiv=df_indiv)
result_df = transformer.to_upload_common(headers=headers)
result_df.to_csv(os.path.join(folder_upload, f"upload_common_sqlite_{today}.csv"), index=False, encoding='utf8')

result_df.head()

# %%
table_name = "공통속성_표준데이터시트"
columns = result_df.columns.to_list()
primary_key = 'SRNo'
update_columns_str = ', '.join([f"{col} = EXCLUDED.{col}" for col in columns if col not in primary_key])

query = f"""INSERT INTO {table_name} ({', '.join(columns)})
    SELECT {', '.join(columns)} FROM temp_table WHERE TRUE
    ON CONFLICT ({primary_key}) 
    DO UPDATE SET 
    {update_columns_str};
    """

# %%
db_path = os.path.join(folder_base, "GSC_INDIV.sqlite")
conn = sqlite3.connect(db_path)
cur = conn.cursor()

# 임시 테이블 생성 및 데이터 삽입
result_df.to_sql('temp_table', conn, if_exists='replace', index=False)

# 기존 테이블에 데이터 병합
cur.execute(query)
# 임시 테이블 삭제
cur.execute("DROP TABLE temp_table")
# 커밋 및 연결 종료
conn.commit()
conn.close()


# try :
#     # 기존 테이블에 데이터 병합
#     cur.execute(query)
#     # 임시 테이블 삭제
#     cur.execute("DROP TABLE temp_table")
#     # 커밋 및 연결 종료
#     conn.commit()
#     conn.close()
#     print("업로드 완료")
# except :
#     # 임시 테이블 삭제
#     cur.execute("DROP TABLE temp_table")
#     conn.close()
#     print("업로드 실패")

# %%



