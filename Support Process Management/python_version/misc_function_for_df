def drop_duplicate_by(df, condition, filter_list, dupl) :
    """ 특정 조건으로 필터링 후 중복 제거
    df, dataframe : 조작할 데이터프레임
    conditoin, str : 필터링할 값이 있는 칼럼 이름
    filter_list, list : 필터링할 값
    dupl, str : 중복 값 검사 대상 항목
    """
    df_filter = df[df[condition].isin(filter_list)]
    df_filter = df_filter.drop_duplicates(subset=[dupl])
    df_filter2 = df[~df[condition].isin(filter_list)]
    df_result = pd.concat([df_filter2, df_filter])
    df_result = df_result.reset_index(drop=True)
    return df_result
