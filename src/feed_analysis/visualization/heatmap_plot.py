


# 1)数据处理部分
import numpy as np
import pandas as pd
def prepare_heatmap_dynamic_feed_col(df, smooth_win=1, lag=1, if_partial=False, start_date=None, end_date=None):

    # 处理时间保证Date列为datetime.date
    df['Date'] = pd.to_datetime(df['Date']).dt.date

    # 删除重复列
    df = df.drop_duplicates(subset=['Date', 'col_num'], keep='first')

    # 如果只处理部分数据
    if if_partial:
        df = df.loc[df['Date'].between(pd.to_datetime(start_date).date(), pd.to_datetime(end_date).date())].copy()
    
    # 显示后续作图的日龄 - 时间映射表
    if 'age' in df.columns:
        date2age = {
            d: age for d, age in zip(df['Date'].values, df['age'].values)
        }
    else:
        date2age = None

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

# 2) 绘图部分
import plotly.express as px
def heatmap_dynamic_feed_col(df_matrix, date2age, unit='单元', color_max_range=80):
    df_pct_plot = df_matrix.astype(float)

    fig = px.imshow(
        df_pct_plot,
        labels=dict(x="时间", y="栏数", color="栏料量变化率%"),
        x=df_pct_plot.columns,           # 列名作为 x 轴
        y=df_pct_plot.index,             # 行索引作为 y 轴

        color_continuous_scale='RdBu_r',
        color_continuous_midpoint=0,       # 0 对应中间颜色（白色）
        range_color=[-color_max_range, color_max_range],

        # text_auto=True,
        aspect="auto",
        title=f"{unit}各栏饲料量变化率"

    )

    fig.update_traces(
        hovertemplate=
        "日期：%{x:%Y-%m-%d}<br>"
        "日龄：" + "%{customdata}<br>"
        "栏号：%{y}<br>"
        "变化率：%{z:.2f}%<extra></extra>",
        customdata=[
            [date2age[d] for d in df_pct_plot.columns if date2age]
        ] * len(df_pct_plot.index)
    )

    # 如果想让 y 轴从上到下显示（行1在最上面）
    fig.update_yaxes(
        autorange="reversed",
        dtick = 1,
        showgrid=False,
        tickfont=dict(size=10)
    )

    tickvals = df_pct_plot.columns       # 每一天的位置
    ticktext = [d.strftime("%Y-%m-%d") for d in tickvals]  # 每一天显示文本
    
    fig.update_xaxes(
        tickvals=tickvals,     # 指定刻度位置（列索引）
        ticktext=ticktext,     # 指定显示文本
        tickformat="%Y-%m-%d",  # 年-月-日
        tickangle=-90,            # 可选：斜显示防止重叠
        tickfont=dict(size=10)
    )
    
    return fig



