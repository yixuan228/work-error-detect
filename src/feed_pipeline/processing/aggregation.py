def split_granularity(df):
    df_col = df[['Date', 'col_num', 'food_col_kg', 'age']].drop_duplicates()
    df_build = df[['Date', 'food_total_kg', 'age']].drop_duplicates()
    return df_col, df_build