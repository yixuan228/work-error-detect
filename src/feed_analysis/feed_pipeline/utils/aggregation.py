# 返回不同精度等级的数据
def split_granularity(df):
    df_col = df[['Date', 'col_num', 'food_col_kg', 'age']].drop_duplicates()
    df_col = df_col.sort_values(by=['Date', 'col_num'], ascending=[True, True])
    df_build = df[['Date', 'food_total_kg', 'age']].drop_duplicates()
    df_build = df_col.sort_values(by=['Date'], ascending=True)
    return df_col, df_build