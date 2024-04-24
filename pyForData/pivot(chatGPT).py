import pandas as pd
"""
data = {
    'SR No': [1, 1, 1, 1, 2, 2, 2],
    'COL': ['A', 'B', 'C', 'D', 'E', 'A', 'B'],
    'VALUE': [2, 34, 51, 3, None, 12, 2]
}
"""

df = pd.DataFrame(data)

# 피벗을 사용하여 데이터 재구성
pivot_df = df.pivot(index='SR No', columns='COL', values='VALUE')

# NaN 값을 NULL로 변환
pivot_df = pivot_df.where(pd.notnull(pivot_df), None)

print(pivot_df)

# melt 함수를 사용하여 원래의 긴 형태로 데이터를 되돌림.
melted_df = pivot_df.reset_index().melt(id_vars='SR No', value_name='VALUE').dropna(subset=['VALUE'])

print(melted_df)

