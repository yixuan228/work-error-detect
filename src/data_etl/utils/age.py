# def calc_relative_age(df, start_date):
#     df['age'] = [(d - start_date).days for d in df['Date']]
#     return df

import pandas as pd
def recalibrate_age(file_path, reference_date:str, reference_age:int):
    """
    更新指定路径的单个parquet文件中的年龄列,并更新保存为parquet
    """
    print(f'已更新保存{file_path}')
    df = pd.read_parquet(file_path, engine='pyarrow')
    df['age'] = [(d - pd.to_datetime(reference_date).date()).days + reference_age for d in df['Date']]
    df.to_parquet(file_path, index=False) 
    # return df
