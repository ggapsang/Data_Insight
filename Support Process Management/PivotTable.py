import pandas as pd

class Table() :
    def __init__(self, df) :
        self.df = df

    def convert_pivot(self) :
        pivot_df = self.df.pivot(index='SRNo', columns='속성명', values='속성값')
        pivot_df.reset_index(inplace=True)
        
        return pivot_df
    
    def melst(self) :
        melted_df = self.df.melt(id_vars='SR No', 
                      var_name='속성명', 
                      value_name='속성값')
        melted_df.dropna(subset=['속성값'], inplace=True)

        return melted_df
