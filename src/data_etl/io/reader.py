import pandas as pd
import numpy as np
from data_etl.config.coding_schema import STD_HEADER_NAME
def read_feed_excel(file_path) -> pd.DataFrame:
    
    return (
        pd.read_excel(file_path, index_col=False, header=3)
        .loc[:, STD_HEADER_NAME.keys()]                             # 选取指定列 
        .rename(columns=STD_HEADER_NAME)                            # 英文列名变更
        .sort_values(by=['Date', 'col_num'])                        # 日期、栏升序
        .drop_duplicates()                                          # 删除重复行
        .assign(Date=lambda x: pd.to_datetime(x["Date"]).dt.date)   # 转为时间格式
        .reset_index(drop=True) 
    )
