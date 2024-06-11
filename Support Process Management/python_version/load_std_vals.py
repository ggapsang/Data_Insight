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
df = pd.read_csv("---.csv", encoding='utf8')
tag_list = df['표준데이터시트'].to_list()

# %%
## step 0 : define the path of the data

base_path = "D:\\(Supporter)\\"
script_path = os.path.join(base_path, "working_file")
db_path = os.path.join(base_path, '개별속성.sqlite3')
conn = sqlite3.connect(db_path)

# %%
table_name = "표준데이터시트_개별속성_240530"
query = f"SELECT * FROM {table_name}"
df = pd.read_sql_query(query, conn)

# %%
df_filtered = df[df['SRNo'].isin(tag_list)]

# %%
df_filtered.to_excel("2201_표준데이터시트_240611.xlsx", index=False)


