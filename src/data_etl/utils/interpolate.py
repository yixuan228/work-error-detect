import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def interpolation(df, index_col='age', method='linear', time_col='Date'):
    """
    根据 index_col 补全缺失行，并对数值列和时间列做线性插值。
    支持 datetime64[ns] 和 datetime.date 类型的时间列。
    会自动对开头和结尾缺失进行外推。
    """
    
    df = df.copy()
    
    # 排序
    df = df.sort_values(by=index_col).reset_index(drop=True)
    
    # 补全 index
    index_full = np.arange(df[index_col].min(), df[index_col].max() + 1)
    df_full = pd.DataFrame({index_col: index_full})
    
    # 合并原始数据
    df_full = df_full.merge(df, on=index_col, how='left')
    
    # 数值列插值（包括开头和结尾缺失）
    num_cols = df_full.select_dtypes(include=[np.number]).columns.tolist()
    df_full[num_cols] = df_full[num_cols].interpolate(method=method, limit_direction='both')
    
    # 时间列插值（按天）
    if time_col in df_full.columns:
        # 确保时间列是 datetime 类型
        df_full[time_col] = pd.to_datetime(df_full[time_col])
        # 按 index_col 线性映射时间
        start_date = df_full[time_col].min()
        end_date = df_full[time_col].max()
        df_full[time_col] = pd.date_range(start=start_date, end=end_date, periods=len(df_full))
    
    return df_full
