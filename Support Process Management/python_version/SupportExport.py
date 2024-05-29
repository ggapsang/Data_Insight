import pandas as pd
import numpy as np
import os
from datetime import datetime
from tqdm import tqdm
import json

class MeltingData() :
    def __init__(self,data_path, dic_cct=None) :
        self.data_path = data_path
        self.dict_cct = dic_cct

    def step1(self) :
        """load data : 데이터 불러오기"""
        try :
            self.df = pd.read_csv(self.data_path, encoding='cp949')
        except :
            self.df = pd.read_csv(self.data_path, encoding='utf8')
        return self
    
    def step2(self, filter_col='속성 그룹 코드', filter_val='03_DATA') :
        """filtering data : 데이터 필터링"""
        self.df_filtered = self.df[self.df[filter_col].isin([filter_val])]
        return self
    
    def step3(self, col_list=None, key=None, sr_col='SR No', cct_col='C|C|T', process_col='공정') :
        """extract the common part in dataframe : 데이터프레임의 공통 속성 부분 추출 : 기본값은 'SR No', '공정', 'C|C|T'"""
        if col_list is None :
            self.df_common = self.df_filtered[[sr_col, process_col, cct_col]]
            self.df_common.drop_duplicates(subset=[sr_col], keep='first', inplace=True)
        else :
            self.df_common = self.df_filtered[col_list]
            self.df_common.drop_duplicates(subset=[key], keep='first', inplace=True)
        
        return self
    
    def step4(self, attr_header=None, drop_null=False, filter_col='속성 그룹 코드', filter_val='03_DATA', key_col='SR No', first_attr_nm='개별속성1', cct_col='C|C|T', value_nm='속성값', var_nm='속성순번') :
        """make attribute dataframe data : 속성값 데이터프레임 생성"""
        self.attr_1_col_no = self.df_filtered.columns.get_loc(first_attr_nm)
        
        if attr_header == None :
            attr_header = self.dict_cct

        self.dict_idx = attr_header['index']
        self.header_list = attr_header['header_list']
        self.count_attrs = attr_header['count_attrs']

        dfs = []
        for k in tqdm(self.dict_idx.keys()) :

            df_sub = self.df_filtered[self.df_filtered[cct_col] == k]
            
            if df_sub.empty !=True :
                df_attr = df_sub[df_sub[filter_col].isin([filter_val])]
                df_attr = df_attr[[key_col] + df_attr.columns[self.attr_1_col_no:].to_list()]
                count_attr = self.count_attrs[k]
                df_attr = df_attr.iloc[:, :count_attr+1]
                dfs.append(df_attr)
        
        results = []
        for df_attr in tqdm(dfs) :
            df_attr = pd.melt(df_attr, id_vars=[key_col], value_vars=df_attr.iloc[:,1:].columns.to_list(), var_name=var_nm, value_name=value_nm, col_level=None, ignore_index=True)
            if drop_null :
                df_attr = df_attr.dropna()
            results.append(df_attr)

        self.df_attrs = pd.concat(results, ignore_index=True)


    def step5(self, key_col='SR No') :
        """merge common dataframe and attribute dataframe : 공통 데이터프레임과 개별속성 데이터프레임 병합"""
        self.df_indiv = pd.merge(self.df_attrs, self.df_common, on=key_col, how='left')
        return self
    
    def step6(self, attr_header=None, var_nm='속성순번', value_nm='속성명', cct_col='C|C|T') :
        """change_attribute_name : 속성명 변경"""
        def change_attribute_name(dict_idx, value_name, cct, header_list) :
            idx = dict_idx[cct]
            dict_attribute_nm = header_list[idx]
            try :
                new_nm = dict_attribute_nm[value_name]
            except :
                new_nm = "Dumb"
            return new_nm
        
        if attr_header == None :
            attr_header = self.dict_cct

        self.dict_idx = attr_header['index']
        self.header_list = attr_header['header_list']
        self.df_indiv[value_nm] = self.df_indiv.progress_apply(lambda x : change_attribute_name(self.dict_idx, x[var_nm], x[cct_col], self.header_list), axis=1)
        return self
    
    def step7(self, value_nm='속성명') :
        """drop dumb value : Dumb 값 제거"""
        self.df_indiv = self.df_indiv[self.df_indiv[value_nm] != 'Dumb']
        return self
    
    def step8(self, key_col='SR_No_ATTR', cat_col='카테고리', class_col='클래스', type_col='타입', value_col='속성값', cct_col='C|C|T', sr_col='SR No', value_nm='속성명') :
        """key 값 생성 및 기타 칼럼 생성"""
        self.df_result = self.df_indiv
        self.df_result = self.df_result[self.df_result[value_nm] != 'Dumb']
        self.df_result[key_col] = self.df_result[sr_col] + '|' + self.df_result[value_nm]
        self.df_result[cat_col] = self.df_result[cct_col].apply(lambda x : x.split("|")[0])
        self.df_result[class_col] = self.df_result[cct_col].apply(lambda x : x.split("|")[1])
        self.df_result[type_col] = self.df_result[cct_col].apply(lambda x : x.split("|")[2])
        self.df_result[value_col] = self.df_result[value_col].apply(lambda x : np.nan if x =='-' else x)

        return self

    def step9(self) :
        """공통 파일 생성"""
        self.df_service = self.df_filtered
        self.df_servie = self.df_service[self.df_service['속성 그룹 코드'].isin(['03_DATA'])]
        self.df_service = self.df_service[['SR No', '파일목록', '공정', '설비번호', '설비카테고리코드', '설비클래스코드', '설비유형코드', 'C|C|T']]
        self.df_service.drop_duplicates(subset=['SR No'], keep='first', inplace=True)

        return self

    def step10(self, key_col='SR_No_ATTR') :
        """중복 key값이 있는 경우 표시함"""
        self.df_duple_key = self.df_result[self.df_result.duplicated(subset=[key_col])]
        return self

    def std_execute(self, show_duplicate=False) :
        self.step1()
        self.step2()
        self.step3()
        self.step4()
        self.step5()
        self.step6()
        self.step7()
        self.step8()
        self.step9()
        self.step10()

        if show_duplicate :
            print(self.df_duple_key)

        return self.df_result
    
    def help(self) :
        print('step1() : load data : 데이터 불러오기') 
        print('step2() : filtering data : 데이터 필터링')
        print('step3() : extract the common part in dataframe : 데이터프레임의 공통 속성 부분 추출')
        print('step4() : make attribute dataframe data : 속성값 데이터프레임 생성')
        print('step5() : merge common dataframe and attribute dataframe : 공통 데이터프레임과 개별속성 데이터프레임 병합')
        print('step6() : change_attribute_name : 속성명 변경')
        print('step7() : drop dumb value : Dumb 값 제거')
        print('step8() : key 값 생성 및 기타 칼럼 생성')
        print('step9() : 공통 파일 생성')
        print('execute() : run all steps')

    def show_attributes(self):
        # 인스턴스 속성
        instance_attributes = self.__dict__
        print("Instance attributes:")
        for attr, value in instance_attributes.items():
            print(f"{attr}")


def get_cct_dictionary(df, df_cct_attrs_count, attr_nm='개별속성1', save_json=True, save_path=os.getcwd()) :
    """속성 해더를 가지고 딕셔너리를 만듬"""

    # 속성명이 있는 행만 추출
    df_headers = df[df['속성 그룹 코드'].isin(['01_속성명'])]
    attr1_col_no = df_headers.columns.get_loc(attr_nm)
    
    df_headers = df_headers[['C|C|T'] + df_headers.columns[attr1_col_no:].to_list()]
    
    # 속성명과 순번에 대한 딕셔너리 생성
    header_list = []
    for i in tqdm(range(len(df_headers))):
        df_header = df_headers.iloc[i].dropna()
        header_list.append(df_header.to_dict())

    # 딕셔너리 생성
    idx_list = [i for i in range(len(header_list))]
    cct_list = [header_list[i]['C|C|T'] for i in range(len(header_list))]
    dict_idx = dict(zip(cct_list, idx_list))
    count_attrs = dict(zip(df_cct_attrs_count['C|C|T'], df_cct_attrs_count['속성 개수']))
    
    dic = {'index' : dict_idx, 'header_list' : header_list, 'count_attrs' : count_attrs}
    # json 파일로 저장
    
    if save_json :
        import json            
        today = datetime.today().strftime("%y%m%d")
        with open(os.path.join(save_path, f'cct_dict_{today}.json'), 'w') as f:
            json.dump(dic, f)

    return dic
