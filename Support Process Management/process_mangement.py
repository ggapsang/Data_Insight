import PivotTables as pt
import pandas as pd

class TableTransformer() :
    
    def __init__(self, df):
        self.df = df
    
    def convert_common_to_indiv(self, df_headers) :
        




        result_df



        return result_df

    def convert_upload_common(self, headers) :
        """공통속성 업로드 포멧으로 데이터를 변환한다"""
        
        result_df = self.df[self.df['속성 그룹 코드']==('03_DATA' or '04_TBD')]
        result_df = result_df[headers]
        try :
            result_df = result_df.rename(columns = {'공정' : '공정번호', '공정별 분류 코드' : '공종별분류코드'})
            result_df = result_df.astype({'공정번호' : str})
        except :
            print("e")
        
        return result_df

    def convert_uplaod_indiv(self, drop_list) :
        """개별속성 업로드 포멧으로 데이터를 변환한다"""
        
        # "2. upload dataform 으로 변환" 부분에서 사용함
        def get_upload_df(df_cct) :
            """하나의 업로드 형식 테이블 완성"""
    
            #### filtered table
            df_header = df_cct[df_cct['속성 그룹 코드']=='01_속성명']
            df_vals = df_cct[df_cct['속성 그룹 코드']=='03_DATA']
            df_header.drop('속성 그룹 코드', axis=1, inplace=True)
            df_vals.drop('속성 그룹 코드', axis=1, inplace=True)
    
            #### 새 해더 이름
            header = df_header.iloc[0].dropna()
            header = header.to_list()
            del header[1:3]
    
            #### 속성값 부분 편집
            df_left = df_vals[['SR No', '공정', '공정별 분류 코드']]
    
            df_right = df_vals[[col for col in df_vals.columns.to_list() if col not in ['공정', '공정별 분류 코드']]]
            head_len = len(header)
            df_right = df_right.iloc[:, :head_len]

            df_right_cols = df_vals.columns.to_list()
            new_nm_cols = dict(zip(df_right_cols, header))

            df_right.rename(columns = new_nm_cols, inplace=True)
    
            tb = Table(df_right)
            pivot_df = tb.melst()
    
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
    
            upload_df = get_upload_df(df_cct)
            upload_dfs.append(upload_df)
    
        ## 통합 업로드 파일
        result_df = pd.concat(upload_dfs, ignore_index=True)
        result_df.rename(columns = {'공정':'공정번호'})
            
        result_df['출처'] = None
        result_df['상태'] = '업로드 대기'
        
        return result_df
    
    def convert_upload_keyin(self) :
        """태그 키인 업로드 포멧으로 데이터를 변환한다"""
        return result_df
    
class UploadValidation() :
        
        def __init__(self, df) :
            self.df = df
        
        def validate_common(self) :
            """공통속성 업로드 데이터를 검증한다"""
            return result_df
        
        def validate_indivi(self) :
            """개별속성 업로드 데이터를 검증한다"""
            return result_df
        
        def validate_keyin(self) :
            """태그 키인 업로드 데이터를 검증한다"""
            return result_df
        
class Reporting() :

    def __init__(self, df) :
        self.df = df

    def report_weekly(self) :
        """주간 리포트를 생성한다"""
        return result_df
    
    def report_different_values(self) :
        """출처별로 차이가 나는 값들을 리포트한다"""
        return result_df



## TEST 공통속성 탬플릿에서 개별속성 템플릿으로 전환

# 1. 불필요한 해더 제거
base_path = "D:\\pypy\\너무해\\"
file = "2201_516_공통작업파일.csv"
f_path = os.path.join(base_path, file)
df_common_working = pd.read_csv(f_path, encoding='utf8')

headers = "D:\\pypy\\너무해\\schema\\info_header.csv"
df_headers = pd.read_csv(headers, encoding='utf8')
df_headers = df_headers.fillna(value=np.nan)


indiv_working_cols = df_headers['개별속성 작업 해더'].to_list()
df_headers['공통속성 작업 해더 매핑'].replace({np.nan: None}, inplace=True)
common_mapping_cols = df_headers['공통속성 작업 해더 매핑'].to_list()

mapping_dic = dict(zip(indiv_working_cols, common_mapping_cols))
mapping_dic

df_indiv_working = pd.DataFrame()
for key, value in mapping_dic.items() :
    if value != None :
        df_indiv_working[key] = df_common_working[value]
    else:
        continue

# 2. 선작업 태그, 표준데이터시트 태그 찾기
df_std_ds =  df_common_working[['SRNo', '대표 SR No', 'cct']][df_common_working['출처']=='2.0표준']
df_std_ds.rename(columns={'SRNo' : '표준데이터시트'}, inplace=True)
df_nonstd_ds = df_common_working[['SRNo', '대표 SR No', 'cct']][df_common_working['출처']=='2.3.비표준시트_수기']
df_nonstd_ds.rename(columns={'SRNo' : '선작업 태그'}, inplace=True)

df_indiv_working = pd.merge(df_indiv_working, df_std_ds, left_on='SR No', right_on='대표 SR No', how='left')
df_indiv_working.drop('대표 SR No', axis=1, inplace=True)
df_indiv_working = pd.merge(df_indiv_working, df_nonstd_ds, left_on='SR No', right_on='대표 SR No', how='left')
df_indiv_working.drop('대표 SR No', axis=1, inplace=True)

series_std = df_indiv_working['cct_x']
series_nonstd = df_indiv_working['cct_y']
df_indiv_working['비교'] = series_std.combine_first(series_nonstd)
df_indiv_working.drop('cct_x', axis=1, inplace=True)
df_indiv_working.drop('cct_y', axis=1, inplace=True)


# 3. CCT 매핑
schema_folder = "D:\\pypy\\너무해\\schema\\"
f_cct = "분류체계.csv"
path_cct = os.path.join(schema_folder, f_cct)
df_cct = pd.read_csv(path_cct)

df_indiv_working = pd.merge(df_indiv_working, df_cct, left_on='타입', right_on='LV6.3_TYPE (DESCRIPTION)')
df_indiv_working.rename(columns={'LV6.1_CATEGORY (DESCRIPTION)' : '카테고리', 'LV6.2_CLASS (DESCRIPTION)' : '클래스', 'LV6.3_TYPE (DESCRIPTION)' : '타입'}, inplace=True)

df_indiv_working['속성 그룹 코드'] = '03_DATA'

def compare_or_nan(row, col_nm_1, col_nm_2) :
    if pd.isna(row[col_nm_1]) or pd.isna(row[col_nm_2]) :
        return np.nan
    else :
        return(row[col_nm_1] == row[col_nm_2] )

df_indiv_working['최종 CCT 변경 유무'] = df_indiv_working.apply(lambda x : compare_or_nan(x, col_nm_1='비교', col_nm_2='C|C|T'), axis=1)

for 3251

df_indiv_working
