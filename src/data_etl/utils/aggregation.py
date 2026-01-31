# 返回不同精度等级的数据
def split_granularity(df):
    df_col = df[['Date', 'col_num', 'food_col_kg']].drop_duplicates(subset=['Date', 'col_num', 'food_col_kg'])
    df_col = df_col.sort_values(by=['Date', 'col_num'], ascending=[True, True])

    df_build = df[['Date', 'food_total_kg']].drop_duplicates(subset=['Date', 'food_total_kg'])
    df_build = df_build.sort_values(by=['Date'], ascending=True)

    # print(df_build[df_build['food_total_kg']>100, :]) # debug

    return df_col, df_build