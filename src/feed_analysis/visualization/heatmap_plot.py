


# 1)数据处理部分
import numpy as np
import pandas as pd
def prepare_heatmap_dynamic_feed_col(df, smooth_win=1, lag=1, if_partial=False, start_date=None, end_date=None):

    # 如果只处理部分数据
    if if_partial:
        df = df.loc[df['Date'].between(pd.to_datetime(start_date).date(), pd.to_datetime(end_date).date())].copy()
    
    # 显示后续作图的日龄 - 时间映射表
    date2age = {
        d: age for d, age in zip(df['Date'], df['age'])
    }

    # 转宽数据计算后续值
    df_wide = df.pivot(index='Date', columns='col_num', values='food_col_kg')
    df_wide.reset_index(inplace=True)
    
    # 构造平滑序列
    df_for_smooth = df_wide.drop(columns=['Date']).replace(0, np.nan)       # 保证NaN不参与平滑计算，临时表格
    df_smooth = df_for_smooth.rolling(window=smooth_win, center=True, min_periods=1).mean()

    # 构造滞后序列
    df_smooth_lag = df_smooth.shift(lag)

    # 计算滞后变化率
    df_pct_change = (df_smooth - df_smooth_lag) / df_smooth_lag * 100 # 转为百分比
    df_pct_change.index = df_wide['Date']       # 添加日期
    df_pct_change = df_pct_change[lag:]         # 剔除空行
    df_matrix = df_pct_change.T                 # 转置，每一行为每一栏，每一列是一个时间节点

    return date2age, df_matrix



