import numpy as np
import pandas as pd
import os
from tqdm import tqdm
from datetime import datetime
import sqlite3

import PivotTable as pt


class TableTransformer() :
    """작업 프로세의 매 단계마다의 템플릿을 이전 템플릿에서 자동으로 변환함"""
    def __init__(self, df, df_cct, df_indiv):
        self.df = df
        self.df_cct = df_cct
        self.df_indiv = df_indiv
    
    def from_common_to_indiv(self, df_headers, df_std_ds_cct, df_nonstd_ds_cct) : #2024.05.02 테스트 완료
        """공통속성 작업 탬플릿에서 개별속성 작업 탬플릿으로 변환한다"""
        # df_headers : 데이터 테이블 스키마 정보를 가지고 있는 해더 파일

        # 1. 데이터프레임 생성 (1) : 해더 매핑으로 값 불러오기
        df_headers['공통속성 작업 해더 매핑'].replace({np.nan: None}, inplace=True)
        headers_indiv = df_headers['개별속성 작업 해더'].to_list()
        headers_common = df_headers['공통속성 작업 해더 매핑'].to_list()

        mapping_indiv_common = dict(zip(headers_indiv, headers_common))

        df_indiv = pd.DataFrame()
        for key, value in mapping_indiv_common.items() :
            if value != None :
                df_indiv[key] = self.df[value]
            else:
                continue
            
        # 2. 데이터프레임 생성(2) : '선작업 태그', '표준데이터시트' 칼럼 값 입력
        df_std_ds = self.df[['SRNo', '대표 SR No', 'cct']][self.df['출처']=='2.0.표준']
        df_std_ds.rename(columns={'SRNo' : '표준데이터시트'}, inplace=True)
        df_nonstd_ds = self.df[['SRNo', '대표 SR No', 'cct']][self.df['출처']=='2.3.비표준시트_수기']
        df_nonstd_ds.rename(columns={'SRNo' : '선작업 태그'}, inplace=True)
        df_std_ds.to_csv("log-표준.csv", encoding='cp949')
        df_nonstd_ds.to_csv("log-비표준.csv", encoding='cp949')

        df_indiv = pd.merge(df_indiv, df_std_ds, left_on='SR No', right_on='대표 SR No', how='left')
        df_indiv.drop('대표 SR No', axis=1, inplace=True)
        df_indiv = pd.merge(df_indiv, df_nonstd_ds, left_on='SR No', right_on='대표 SR No', how='left')
        df_indiv.drop('대표 SR No', axis=1, inplace=True)

        # 2-1. CCT 비교(표준 데이터시트는 별도의 파일에 불러 와야 함)        
        df_std_ds_cct = df_std_ds_cct[['SR No', 'C|C|T']]
        df_nonstd_ds_cct = df_nonstd_ds_cct[['New SR No', 'CCT']]

        df_indiv = pd.merge(df_indiv, df_std_ds_cct, left_on='표준데이터시트', right_on='SR No', how='left')
        df_indiv = pd.merge(df_indiv, df_nonstd_ds_cct, left_on='선작업 태그', right_on='New SR No', how='left')

        # series_std = df_std_ds_cct['C|C|T']
        # series_nonstd = df_indiv['CCT']
        # df_indiv.to_csv("log_df_indiv.csv", encoding='cp949')
        df_indiv['비교'] = df_indiv['C|C|T'].combine_first(df_indiv['CCT'])
        # print(df_indiv)

        df_indiv.drop('C|C|T', axis=1, inplace=True)
        df_indiv.drop('CCT', axis=1, inplace=True)
        df_indiv.drop('SR No_y', axis=1, inplace=True)
        df_indiv.rename(columns={'SR No_x' : 'SR No'}, inplace=True)

        # 3. 데이터프레임 생성(3) : CCT 매핑 : 작업템플릿 + CCT(분류체계) LEFT JOIN
        df_indiv = pd.merge(df_indiv, self.df_cct, left_on='타입', right_on='LV6.3_TYPE (DESCRIPTION)', how='left')
        df_indiv.drop(columns='LV6.3_TYPE (DESCRIPTION)')
        df_indiv.rename(columns={'LV6.1_CATEGORY (DESCRIPTION)' : '카테고리', 'LV6.2_CLASS (DESCRIPTION)' : '클래스'}, inplace=True)

        # 4. 데이터프레임 생성(4) : '작업자', '속성 그룹 코드', '최종 CCT 변경 유무', '비고' 칼럼 추가
        df_indiv['속성 그룹 코드'] = '03_DATA'
        df_indiv['작업자'] = None
        df_indiv['비고'] = None

        def compare_or_nan(row, col_nm_1, col_nm_2) :
            """두 칼럼의 값이 같은지 비교하고, 둘 중 하나라도 NaN이면 NaN을 반환한다"""
            if pd.isna(row[col_nm_1]) or pd.isna(row[col_nm_2]) :
                return np.nan
            else :
                return(row[col_nm_1] != row[col_nm_2] )

        df_indiv['최종 CCT 변경 유무'] = df_indiv.apply(lambda x : compare_or_nan(x, col_nm_1='비교', col_nm_2='C|C|T'), axis=1)

        # 5. 데이터프레임 생성(5) : 속성 해더와 결합
        attr_cols = [f"속성{i}" for i in range(1, 3252)]
        df_attrs = pd.DataFrame(columns=attr_cols)

        df_indiv = pd.concat([df_indiv, df_attrs], axis=1) # 속성 해더와 결합

        # 6. 데이터프레임 생성(6) : 'MDM 반영 여부'에 N인 값들은 제외
        df_mdm_upload = self.df.loc[(self.df['MDM 반영 여부'] == 'Y') | (self.df['MDM 반영 여부'] == 'Y(배관)'), 'SRNo']

        # print(df_indiv.columns.to_list())
        df_indiv = df_indiv[df_indiv['SR No'].isin(df_mdm_upload)]

        # 7. df_common과 한번 더 left join
        df_common_cct = self.df[['SRNo', 'CATEGORY', 'CLASS', 'TYPE']]
        df_indiv = pd.merge(df_indiv, df_common_cct, left_on='SR No', right_on='SRNo', how='left')
        df_indiv.drop('카테고리', axis=1, inplace=True)
        df_indiv.drop('클래스', axis=1, inplace=True)
        df_indiv.drop('타입', axis=1, inplace=True)
        df_indiv.drop('SRNo', axis=1, inplace=True)

        df_indiv.rename(columns={'CATEGORY' : '카테고리', 'CLASS' : '클래스', 'TYPE' : '타입'}, inplace=True)

        # 8. 2에서 left join으로 1:n 매핑된 값들중 하나만 남기고 제거
        df_indiv.drop_duplicates('SR No', inplace=True)

        # 9. 해더 열 순서 정렬
        order_list = df_headers['개별속성 작업 해더'].to_list()

        df_indiv = df_indiv[order_list]

        return df_indiv

    def to_upload_common(self, headers : list) :
        """개별속성 작업 템플릿에서 공통속성을 업로드할 포멧으로 데이터를 변환한다"""
        
        result_df = self.df[self.df['속성 그룹 코드']==('03_DATA' or '04_TBD')]
        result_df = result_df[headers]

        result_df.rename(columns = {'공정' : '공정번호', 'SR No' : 'SRNo', 'Tag No' : 'TagNo', 'Tag No 수정' : 'TagNo수정', '공정별 분류 코드' : '공종별분류코드'}, inplace=True)
        result_df = result_df.astype({'공정번호' : str})
        
        result_df.drop('속성 그룹 코드', axis=1, inplace=True)

        return result_df

    def to_upload_indiv(self, drop_list) : #20240503 테스트 완료
        """개별속성 작업 템플릿에서 개별속성을 업로드할 포멧으로 데이터를 변환한다"""
        
        # "2. upload dataform 으로 변환" 부분에서 사용함
        def get_upload_single_df(df_cct) :
            """하나의 업로드 형식 테이블 완성"""
    
            #### filtered table
            df_header = df_cct[df_cct['속성 그룹 코드']=='01_속성명']
            df_vals = df_cct[df_cct['속성 그룹 코드']=='03_DATA']
            df_cct.drop('속성 그룹 코드', axis=1, inplace=True)
            df_header.drop('속성 그룹 코드', axis=1, inplace=True)
            df_vals.drop('속성 그룹 코드', axis=1, inplace=True)

    
            #### 해더 이름 새로 매핑
            header = df_header.iloc[0].dropna()
            header = header.to_list()
            head_len = len(header)
            df_header_common = df_cct.iloc[:, :head_len]
            header_common = df_header_common.columns.to_list()
            new_nm_cols = dict(zip(header_common, header))
    

            #### 속성값 부분 편집
            df_left = df_vals[['SR No', '공정', '공정별 분류 코드']]
    
            df_right = df_vals[[col for col in df_vals.columns.to_list() if col not in ['공정', '공정별 분류 코드', '출처']]] #출처 추가
            head_len = len(header)
            df_right = df_right.iloc[:, :head_len]

            df_right.rename(columns = new_nm_cols, inplace=True)
    
            pivot_df =df_right.melt(id_vars='SR No', var_name='속성명', value_name='속성값', ignore_index=False)
            pivot_df = pivot_df.dropna(subset=['속성값'])
    
            upload_df = pd.merge(left=pivot_df, right=df_left, how='left', on='SR No')
            upload_df = upload_df[['공정', 'SR No', '공정별 분류 코드', '속성명', '속성값']]# 칼럼 순서 변경

            return upload_df
        
        
        # 1. 불필요한 칼럼 제거
        filtered_columns = [col for col in self.df.columns.to_list() if col not in drop_list]
        df_2 = self.df[filtered_columns]
        
        # 2. upload data form으로 변환
        cct_codes = df_2['공정별 분류 코드'].unique()
        
        ## cct별 리스트 만들기
        upload_dfs = []
        for cct_code in tqdm(cct_codes) :
    
            df_cct = df_2[df_2['공정별 분류 코드']==cct_code]

            try :
                upload_df = get_upload_single_df(df_cct)
            except :
                print(cct_code)
            upload_dfs.append(upload_df)
    
        ## 통합 업로드 파일
        result_df = pd.concat(upload_dfs, ignore_index=True)
        result_df.rename(columns = {'공정':'공정번호', '공정별 분류 코드':'공종별분류코드'}, inplace=True)
            
        result_df['출처'] = None
        result_df['상태'] = '업로드 대기'

        result_df = result_df[result_df['공종별분류코드'] != result_df['속성명']]
        
        return result_df

class InsertIndiv() :
    """개별속성 작업 템플릿으로 변환한 곳에다 선작업한 속성값들을 우선순위에 따라 반영한다"""
    def __init__(self, conn, query, df_indiv) :
        self.conn = conn
        self.query = query
        
        self.attrs_df = pd.read_sql_query(self.query, self.conn)
        self.df_indiv = df_indiv

    def get_df(self) :
        return self.attrs_df
    
    def insert_primary_SrNo(self) :
        """개별속성 작업 템플릿에 있는 SR No를 대표 키로 가져온다"""
        df_primary_key = pd.DataFrame()
        df_primary_key['대표'] = self.df_indiv['SR No']

        series_std_key = self.df_indiv['표준데이터시트']
        series_nonstd_key = self.df_indiv['선작업 태그']
        series_combine = series_std_key.combine_first(series_nonstd_key)

        df_primary_key['SR No'] = series_combine

        self.attrs_df = pd.merge(self.attrs_df, df_primary_key, on='SR No', how='left')

        return self.attrs_df
    
    def combine_by_priority(self, do_insert_primary_SrNo=False) :
        """우선순위에 따라 속성값을 모아둔 시리즈를 만들고 우선순위에서 밀려난 항목들은 제거한다"""

        if do_insert_primary_SrNo :
            self.attrs_df = self.insert_primary_SrNo()

        attr_std = self.attrs_df[self.attrs_df['출처']=='표준데이터시트']
        attr_nonstd = self.attrs_df[self.attrs_df['출처']=='선작업 태그']
        attr_eleDB = self.attrs_df[self.attrs_df['출처']=='전기설비DB']
        attr_eleLoad = self.attrs_df[self.attrs_df['출처']=='LoadList']
        attr_psm = self.attrs_df[self.attrs_df['출처']=='PSM']
        
        attr_list = [attr_eleDB, attr_std, attr_nonstd, attr_eleLoad, attr_psm]

        self.attrs_df = pd.concat(attr_list, ignore_index=True)

        result_df = pd.DataFrame()

        for key, group in self.attrs_df.groupby('대표') :
            valid_row = group.dropna().head(1)
            if not valid_row.empty :
                result_df = pd.concat([result_df, valid_row], ignore_index=True)
            else :
                result_df = pd.concat([result_df, group.head(1)], ignore_index=True)

        return self.attrs_df
    
    def melt(self, df_attr_headers, do_combine_by_priority=False) :
        
        if do_combine_by_priority :
            self.attrs_df = self.combine_by_priority()

        attrs_df_cct_list = []
        for index, row in df_attr_headers.iterrows() :
            headers = row.values
            df_attr_cct = pd.DataFrame(columns=headers)
            attrs_df_cct = self.attrs_df[self.attrs_df['공정별 분류 코드']==df_attr_headers['공정별 분류 코드']]

            pt_tb = pt.Table(attrs_df_cct)
            melt_df_cct = pt_tb.melst()
            df_attr_cct = pd.concat([df_attr_cct, melt_df_cct], ignore_index=True)
            
            attrs_df_cct_list.append(df_attr_cct)

        result_df = pd.DataFrame(columns=df_attr_headers.columns)

        for df in attrs_df_cct_list :
            cols_mapping = dict(zip(df.columns, df_attr_headers.columns))
            df = df.rename(columns = cols_mapping)

            result_df = pd.concat([result_df, df], ignore_index=True)

        return result_df


"""미구현"""
class UploadValidation() :
        
        def __init__(self, df) :
            self.df = df
        
        def validate_common(self) :
            """공통속성 업로드 데이터를 검증한다"""
            return 0
        
        def validate_indivi(self) :
            """개별속성 업로드 데이터를 검증한다"""
            return 0
        
        def validate_keyin(self) :
            """태그 키인 업로드 데이터를 검증한다"""
            return 0
        
class Reporting() :

    def __init__(self, df) :
        self.df = df

    def report_weekly(self) :
        """주간 리포트를 생성한다"""
        return 0
    
    def report_different_values(self) :
        """출처별로 차이가 나는 값들을 리포트한다"""
        return 0
