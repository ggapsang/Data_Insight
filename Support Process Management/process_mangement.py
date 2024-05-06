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
    
    def help(self) :
        print("from_common_to_indiv : 공통속성 작업 탬플릿에서 개별속성 작업 탬플립으로 변환")
        print("to_upload_common : 개별속성 작업 템플릿에서 공통속성을 업로드할 포멧으로 데이터를 변환")
        print("to_upload_indiv : 개별속성 작업 템플릿에서 개별속성을 업로드할 포멧으로 데이터를 변환")
        print("help : 도움말 출력")

class ReadDB() :
    """SQLite DB에서 데이터를 읽어와 데이터프레임으로 변환한다"""

    def __init__(self, db_path) :
        self.db_path = db_path

    def read_db_to_dataframe(self, table_name) :
        """DB에서 데이터를 불러와 데이터프레임으로 변환한다"""
        conn = sqlite3.connect(self.db_path)
        query = f"SELECT * FROM {table_name}"
        df = pd.read_sql_query(query, conn)
        conn.close()

        return df

class InsertAttrstPreprocessing() :
    """개별속성 작업 템플릿에 데이터를 입력하기 전에 전처리를 수행한다"""
    
    def __init__(self, df_working_f, df_attrs_headers) :
        self.df_working_f = df_working_f
        self.df_attrs_headers = df_attrs_headers[df_attrs_headers['속성 그룹 코드']=='01_속성명']

    def step0_1(self) :
        """"df_working_f에서 ['표준데이터시트', '선작업데이터'] 칼럼의 값을 보충한다"""

        self.df_working_f['표준데이터시트'] = self.df_working_f.apply(lambda x : x['SRNo'] if x['표준데이터시트'] != np.nan and x['출처'] =='표준' else x['표준데이터시트'], axis=1)

        self.df_working_f['선작업데이터'] = self.df_working_f.apply(lambda x : x['SRNo'] if x['선작업데이터'] != np.nan and x['출처'] =='비표준' else x['선작업데이터'], axis=1)

        return self
    
    def step0_2(self, df_attrs, col_name='표준데이터시트') :
        """개별속성 테이블에 대표 SRNo를 추가한다"""

        def get_representative_srno(df_base, df_right, lookup_columns) :
            df_right_join = df_right[['SRNo', lookup_columns]]
            df_merge = pd.merge(df_base, df_right_join, how='left', left_on='SRNo', right_on=lookup_columns)
            df_merge.drop(lookup_columns, axis=1, inplace=True)
            df_merge.rename(columns={'SRNo_x' : 'SRNo', 'SRNo_y':'대표 SRNo'}, inplace=True)
            return df_merge

        self.df_attrs = df_attrs

        self.df_attrs = get_representative_srno(self.df_attrs, self.df_working_f, col_name)
        self.df_attrs.dropna(subset=['대표 SRNo'], inplace=True)

        return self

class InsertAttrsPipeline() :
    """sqlite db에서 가져와 df로 바꾼 데이터들을 개별속성 작업 템플릿에 입력"""

    def __init__(self, df_working_f, df_attrs, df_attrs_headers) :
        self.df_working_f = df_working_f
        self.df_attrs_headers = df_attrs_headers
        self.df_attrs = df_attrs

    def loop_step1(self, i) :
        """해더 파일에서 공종별 분류 코드를 하나 가져온다"""

        self.type_code = self.df_attrs_headers['공종별 분류 코드'][i]

        # df_working_filtered
        self.df_working_filtered = self.df_working_f[self.df_working_f['속성 그룹 코드']==self.type_code] == '03_DATA'
        self.df_working_filtered = self.df_working_f[self.df_working_f['공종별 분류 코드']==self.type_code]

        attrs_columns = self.df_attrs_headers.iloc[i].to_list()

        self.df_cct = pd.DataFrame(columns=attrs_columns)
        self.df_cct = self.df_cct.loc[:, self.df_cct.columns.notnull()]

        return self
    
    def loop_step2(self) :
        """개별속성 테이블 피벗 준비"""

        self.df_attrs_filtered = self.df_attrs[['대표 SRNo', '속성명', '속성값', '공종별분류코드']]
        self.df_attrs_filtered = self.df_attrs_filtered[self.df_attrs_filtered['공종별분류코드']==self.type_code]

        self.df_attrs_filtered.rename(columns={'대표 SRNo' : 'SRNo'}, inplace=True)
        self.df_attrs_filtered.drop('공종별분류코드', axis=1, inplace=True)

        return self
    
    def loop_step3(self) :
        """개별속성 테이블 피벗팅"""

        tb = pt.Table(self.df_attrs_filtered)
        self.df_tb = tb.convert_pivot()

        return self

    def loop_step4(self) :
        """빈 df_cct 데이터프레임과 합침"""
        
        self.df_cct_result = pd.DataFrame()
        self.df_cct_result = pd.concat([self.df_cct, self.df_tb], axis=0, ignore_index=True)

        return self
    
    def loop_step5(self) :
        """피벗팅하여 입력한 속성값과 df_working_f에 있는 정보들을 합침"""
        common_columns = self.df_cct_result.columns.intersection(self.df_attrs_headers.columns)
        self.df_cct_result = pd.merge(self.df_cct_result, self.df_working_filtered[common_columns], how='left', on='SRNo')
        
        for col in self.df_cct_result.columns :
            if '_x' in str(col) :
                self.df_cct_result.drop(col, axis=1, inplace=True)
            elif '_y' in str(col) :
                self.df_cct_result.rename(columns={col : col.replace('_y', '')}, inplace=True)
        
        # 열 순서 정렬
        self.df_cct_result = self.df_cct_result[self.df_cct.columns.to_list()]

        return self
    
    def loop_step6(self) :
        """공종별 분류 코드, 속성 그룹 코드 추가"""

        self.df_cct_result[self.type_code] = self.type_code
        self.df_cct_result['01_속성명'] = '03_DATA'

        return self
    
    def loop_step7(self) :
        """해더 정리"""

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
        
        self.df_cct_result = make_header_to_firts_row(self.df_cct_result)
        main_header = cut_header(self.df_working_filtered, self.df_cct_result)

        self.df_cct_result = change_header(self.df_cct_result, main_header)

        return self
    
    def excute(self, i) :
        """실행"""
        self.loop_step1(i)
        self.loop_step2()
        self.loop_step3()
        self.loop_step4()
        self.loop_step5()
        self.loop_step6()
        self.loop_step7()

        return self.df_cct_result