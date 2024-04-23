"""examlple
SR No, Col, Value
1, A, 2
1, B, 34
1, C, 51
1, D, 3
2, E
2, A, 12
2, B, 2
"""

# SQLite pivot Table
"""
CREATE VIEW pivoted_table AS
SELECT 
  SR_No,
  MAX(CASE WHEN COL = 'A' THEN VALUE ELSE NULL END) AS A,
  MAX(CASE WHEN COL = 'B' THEN VALUE ELSE NULL END) AS B,
  MAX(CASE WHEN COL = 'C' THEN VALUE ELSE NULL END) AS C,
  MAX(CASE WHEN COL = 'D' THEN VALUE ELSE NULL END) AS D,
  MAX(CASE WHEN COL = 'E' THEN VALUE ELSE NULL END) AS E
FROM 
  your_table_name
GROUP BY 
  SR_No;
"""

"""
여기서 `MAX` 함수는 단순히 각 `SR_No`와 `COL` 조합에서 단일 값만 존재한다고 가정할 때 사용됩니다. 
각 `SR_No`에 대해 해당하는 `COL` 값이 하나만 있을 때, `MAX` (또는 `MIN`) 함수를 사용해 그 값을 가져오는 것입니다. 
만약 `COL`별로 여러 값이 존재한다면, `MAX`는 그 중 최댓값만을 반환할 것입니다.

`SR_No`에 대해 각 `COL`에 하나의 값만 있고, 단순히 형식을 변환하는 경우에는 `MAX` 함수가 올바른 결과를 반환합니다. 
예시에서는 각 `SR_No`에 대해 각 `COL`이 유니크하다고 가정하고 있습니다. 
그러나 여러 값이 있는 경우, 이 방식은 부적절하고 다른 접근 방식이 필요합니다.

MAX 함수는 텍스트 값에도 작동하지만, 숫자처럼 직관적인 "최대값"을 반환하지 않습니다.
텍스트 값의 경우 MAX 함수는 알파벳 순으로 가장 뒤에 오는 값을 반환합니다. 
이 경우에는 SR_No별로 COL 값이 유일하다고 가정했기 때문에, MAX 함수를 사용하여 해당 COL의 텍스트 값을 가져오는 것입니다.

여기서 MAX 함수는 단순히 그룹화된 결과 집합에서 각 COL에 대한 단일 값을 선택하는데 사용됩니다. 
그룹화된 각 집합 내에서 COL이 유일하다면 MAX는 해당 COL의 유일한 값을 반환하게 됩니다.

그러므로, 만약 COL 값이 텍스트이고 각 SR_No에 대해 유일한 값이 보장된다면 위의 쿼리는 올바르게 작동하여 각 SR_No에 대해 모든 COL 값을 해당하는 행에 나열하게 됩니다.
"""

import pandas as pd

# 데이터프레임을 만듭니다. 이 예제에서는 데이터가 이미 있는 것으로 가정합니다.
data = {
    'SR No': [1, 1, 1, 1, 2, 2, 2],
    'COL': ['A', 'B', 'C', 'D', 'E', 'A', 'B'],
    'VALUE': [2, 34, 51, 3, None, 12, 2]
}

df = pd.DataFrame(data)

# 피벗을 사용하여 데이터를 재구성합니다.
pivot_df = df.pivot(index='SR No', columns='COL', values='VALUE')

# NaN 값을 NULL로 변환하고 싶다면 다음을 수행합니다.
pivot_df = pivot_df.where(pd.notnull(pivot_df), None)

print(pivot_df)

# pivot_df는 이미 피벗된 데이터프레임이라고 가정합니다.

# melt 함수를 사용하여 원래의 긴 형태로 데이터를 되돌립니다.
melted_df = pivot_df.reset_index().melt(id_vars='SR No', value_name='VALUE').dropna(subset=['VALUE'])

print(melted_df)

