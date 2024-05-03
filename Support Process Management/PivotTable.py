import pandas as pd

class Table() :
    def __init__(self, df) :
        self.df = df

    def convert(self) :
        pivot_df = self.df.pivot(index='SR No', columns='COL', values='VALUE') 
        
        return pivot_df
    
    def melst(self) :
        melted_df = self.df.melt(id_vars='SR No', 
                      var_name='속성명', 
                      value_name='속성값')
        melted_df.dropna(subset=['속성값'], inplace=True)

        return melted_df
