# 全局数据处理配置
SMOOTH_WIN = 1  # 平滑窗口大小
LAG = 1         # 滞后阶数

import numpy as np
import pandas as pd

# 全局绘图参数配置
from matplotlib import font_manager
CN_font = font_manager.FontProperties(fname='C:/Windows/Fonts/simhei.ttf')
# 1. 处理食料数据

def prepare_feed_data_2heatmap(df, window=SMOOTH_WIN, lag=LAG, if_partial=False, start_date=None, end_date=None):
    # 数据读取与平滑
    # Food data
    # df = pd.read_excel(f'data/{file_name}.xlsx', sheet_name=sheet_name, index_col=False).rename(columns={'日期': 'Date'}).drop(columns=['Age'], errors='ignore')
    # df['Date'] = pd.to_datetime(df['Date'])

    # 如果只处理部分数据
    if if_partial:
        df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)].reset_index(drop=True)

    # 构造平滑序列
    df_for_smooth = df.drop(columns=['Date']).replace(0, np.nan)   # 保证NaN不参与平滑计算，临时表格

    df_smooth = (
        df_for_smooth.rolling(window=window, center=True, min_periods=1).mean()
    )
    df_smooth[df == 0] = 0  # 回填0

    # 构造滞后序列
    df_smooth_lag = df_smooth.shift(lag)

    # 计算滞后变化率
    df_pct_change = (df_smooth - df_smooth_lag) / df_smooth_lag * 100 # 转为百分比
    df_pct_change.index = df['Date']    # 添加日期
    df_pct_change = df_pct_change[lag:]   # 剔除空行

    # 转标准数字，保证顺利绘图
    df_num = df_pct_change.apply(pd.to_numeric, errors='coerce')
    df_num = df_num.fillna(float('inf'))
    df_num.index = df_pct_change.index

    df_num = df_num.T

    return df_num, df_pct_change

def heatmap_dynamic_feed_col(df_num, df_pct_change):
    df_pct_plot = df_num.astype(float)

    import plotly.express as px

    fig = px.imshow(
        df_pct_plot,
        labels=dict(x="时间", y="栏数", color="料量变化百分比/%"),
        x=df_pct_plot.columns,           # 列名作为 x 轴
        y=df_pct_plot.index,             # 行索引作为 y 轴

        color_continuous_scale='RdBu_r',
        color_continuous_midpoint=0,       # 0 对应中间颜色（白色）
        range_color=[-80, 80],

        # text_auto=True,
        aspect="auto",
        title="3栋1单元各栏饲料量变化率-%"

    )

    # 如果想让 y 轴从上到下显示（行1在最上面）
    fig.update_yaxes(
        autorange="reversed",
        dtick = 1,
        showgrid=False,
        tickfont=dict(size=10)
    )

    # tickvals = list(range(0, len(df_pct_plot.columns), 3))   # 列索引位置
    # ticktext = [df_pct_plot.columns[i].strftime("%Y-%m-%d") for i in tickvals]
    fig.update_xaxes(
        # tickvals=tickvals,     # 指定刻度位置（列索引）
        # ticktext=ticktext,     # 指定显示文本
        tickformat="%Y-%m-%d",  # 年-月-日
        tickangle=-45,            # 可选：斜显示防止重叠
        tickfont=dict(size=10)
    )
    
    return fig
    fig.write_html('figure/3栋1单元各栏饲料量变化率热力图.html', include_plotlyjs='cdn', full_html=True)
    fig.show()