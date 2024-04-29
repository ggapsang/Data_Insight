import PivotTables as pt
import pandas as pd

class TableTransformer() :
    
    def __init__(self, df):
        self.df = df
    
    def convert_upload_common(self) :
        """공통속성 업로드 포멧으로 데이터를 변환한다"""
        return result_df

    def convert_uplaod_indivi(self) :
        """개별속성 업로드 포멧으로 데이터를 변환한다"""
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
