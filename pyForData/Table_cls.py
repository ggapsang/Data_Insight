import pandas as pd
"""
data = {
    'SR No': [1, 1, 1, 1, 2, 2, 2],
    'COL': ['A', 'B', 'C', 'D', 'E', 'A', 'B'],
    'VALUE': [2, 34, 51, 3, None, 12, 2]
}
"""

class Table() :
    def __init__(self, df) :
        self.df = df

    def convert(self) :
        pivot_df = df.pivot(index='SR No', columns='COL', values='VALUE') 
        
        return pivot_df
    
    def melst(self) :
        melted_df = df.reset_index().melt(id_vars='SR No', value_name='VALUE').dropna(subset=['VALUE'])
        return melted_df
