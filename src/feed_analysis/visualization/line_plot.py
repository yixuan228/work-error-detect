from matplotlib import font_manager
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd

from feed_analysis.config.path import PATH_FIGURE_PNG, PATH_FIGURE_HTML
CN_font = font_manager.FontProperties(fname='C:/Windows/Fonts/simhei.ttf')

# 静态 单个 单元饲料总量变化折线图
def line_static_feed_build(df, unit:str, if_partial=False, start_date=None, end_date=None, save_img=False):
    """"
    静态单个单元层饲料总量变化折线图绘制, 调用plt包
    """

    if if_partial:
        df = df.loc[df['Date'].between(pd.to_datetime(start_date).date(), pd.to_datetime(end_date).date()), :]

    plt.figure(figsize=(12, 5))
    plt.plot(df['Date'], df['food_total_kg'], label='单元总喂食量', color='red')

    plt.xlabel('日期', fontproperties=CN_font)
    plt.ylabel('饲喂量/kg', fontproperties=CN_font)
    plt.title(f'{unit}饲喂料总量', fontproperties=CN_font)

    # 图细节
    ax = plt.gca()
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    ax.grid(axis='x', linestyle='--', color='gray', alpha=0.4)  # 只显示 x 轴网格线

    plt.xticks(rotation=90, fontproperties=CN_font, fontsize=7)
    plt.legend(prop=CN_font)
    plt.tight_layout()
    plt.savefig(PATH_FIGURE_PNG/f'{unit}-饲料总量变化.png', dpi=300, bbox_inches='tight') if save_img else None
    plt.show()

import pandas as pd
import plotly.graph_objects as go

def line_dynamic_single(df, y: str, plot_title:str, color:str='red', if_partial=False, start_date=None, end_date=None,
):
    """
    绘制单条时序折线图，x为日期，y为指定列，支持部分日期范围。
    """

    df_plot = df.copy()

    # 时间裁剪
    if if_partial:
        df_plot = df_plot.loc[df_plot['Date'].between(pd.to_datetime(start_date).date(), pd.to_datetime(end_date).date())]

    fig = go.Figure()

    # 折线
    fig.add_trace(
        go.Scatter(
            x=df_plot['Date'],
            y=df_plot[y],
            mode='lines',
            name='喂食量',
            line=dict(color=color, width=2),
            marker=dict(size=4),
            customdata=df_plot['age'],
            hovertemplate=
            '日期: %{x}<br>'
            '日龄: %{customdata} 天<br>'
            '喂食量: %{y} kg'
        )
    )

    fig.update_layout(
        title=f'{plot_title}时序图',
        xaxis_title='日期',
        yaxis_title='喂食量(kg)',
        template='plotly_white',
    )

    # x 轴日期格式 & 网格
    fig.update_xaxes(
        tickformat='%Y-%m-%d',
        tickangle=-90,
        dtick='D1',              # 每天一个刻度
        showgrid=True,
    )

    return fig

# 动态 多单元 饲料总量变化折线图
import plotly.express as px
def line_dynamic_feed_build(df_list:list, name_list:list, if_partial=False, start_date=None, end_date=None, save_img=False):
    """
    绘制多单元饲料总量变化折线图（包含均值），支持部分日期范围和导出 HTML。
    """

    ## ------ 读取数据部分 ------
    # 设置第一个表格（先根据日期和采食量去重，再设置日期为索引）
    df = df_list[0].loc[:, ['Date', 'food_total_kg']]

    age_df = df_list[0].loc[:, ['Date', 'age']]     # 取日龄列
    age_df['Date'] = pd.to_datetime(age_df['Date'])

    df.set_index('Date', inplace=True)
    
    for i in range(1, len(df_list)):
        df_temp = df_list[i].copy()
        df_temp.set_index('Date', inplace=True)     # 设置日期为索引,方便对齐
        df[i] = df_temp['food_total_kg']
    # 更名
    df.columns = name_list

    # 计算平均值
    df['均值'] = df.mean(axis=1)
    # 日期格式化
    df.index = pd.to_datetime(df.index)

    # 筛选数据至合适日期
    if if_partial:
        df = df[(df.index >= pd.to_datetime(start_date)) & (df.index <= pd.to_datetime(end_date))]

    ## ------ 数据转换部分 宽数据转长数据 ------
    df_long = df.reset_index().melt(id_vars='Date', var_name='Unit', value_name='Food Intake (Kg)')
    
    # 在长数据后添加日龄
    df_long['age'] = df_long['Date'].map(age_df.set_index('Date')['age'])

    ## ------ 绘图部分 ------
    # 交互图像
    # 绘制交互曲线
    fig = px.line(df_long, 
                x='Date', y='Food Intake (Kg)', 
                color='Unit', 
                title=f'各单元总料量变化趋势曲线', 
                custom_data=['age'],
                )

    # 突出显示均值曲线
    highlight_unit = '均值'
    fig.update_traces(
        selector=lambda trace: trace.name == highlight_unit,
        line=dict(width=3, color='red'),    # 加粗
        marker=dict(size=8),                # 增大标记点
        opacity=0.6,                         # 提高不透明度
        hovertemplate=
        "<b>单元</b>: %{fullData.name}<br>"
        "<b>日龄</b>: %{customdata[0]} 天<br>"
        "<b>日期</b>: %{x:%Y-%m-%d}<br>"
        "<b>饲喂量</b>: %{y:.2f} kg"
        "<extra></extra>"
    )

    # 更新其余曲线
    fig.update_traces(
        selector=lambda trace: trace.name != highlight_unit,
        line=dict(width=1),
        opacity=0.6,
        hovertemplate=
        "<b>单元</b>: %{fullData.name}<br>"
        "<b>日龄</b>: %{customdata[0]} 天<br>"
        "<b>日期</b>: %{x:%Y-%m-%d}<br>"
        "<b>饲喂量</b>: %{y:.2f} kg"
        "<extra></extra>"
    )

    # 设置画布大小和轴标签
    fig.update_layout(
        autosize=True,       # 自动调整画布大小
        height=500,          # 高度可以固定，也可以尝试设置为 None 自适应
        xaxis_title='日期',
        yaxis_title='料量/kg')

    # 设置 x 轴
    fig.update_xaxes(
        dtick="D1",        # 每天一个刻度
        tickangle=-90,      # 标签旋转90度
        tickformat="%m-%d", # 可选：只显示月-日
        tickfont=dict(size=7)  # 调小字体，例如10号
    )

    if save_img:
        fig.write_html(PATH_FIGURE_HTML / f'三栋1-4单元总料量变化趋势曲线.html', include_plotlyjs='cdn', full_html=True)

    return fig


# 动态 单单元
import plotly.express as px
def line_dynamic_feed_column(df, unit:str, if_partial=False, start_date=None, end_date=None, save_img=False):
    """
    动态单个单元 栏位级饲料总量（包含平均值）变化折线图绘制，支持部分日期范围和导出 HTML。
    """ 

    if if_partial:
        df = df.loc[df['Date'].between(pd.to_datetime(start_date).date(), pd.to_datetime(end_date).date()), :]

    # 转宽表矩阵，添加Total(均值)
    df_age = df[['Date', 'age']].drop_duplicates()  # 获取日期和年龄
    df = df[['Date', 'col_num', 'food_col_kg']]
    df_wide = df.pivot(index='Date', columns='col_num', values='food_col_kg')
    df_wide['均值'] = df_wide.mean(axis=1)
    df_long = df_wide.reset_index().melt(id_vars='Date', var_name='Column', value_name='value')

    # 长数据添加日龄项
    df_long['age'] = df_long['Date'].map(df_age.set_index('Date')['age'])
    
    fig = px.line(df_long, 
                  x='Date', y='value', color='Column', 
                  title=f'{unit} 28列单栏喂料量变化曲线',
                  custom_data=['age'],
                  )
    
    fig.update_layout(
        autosize=True,       # 自动调整画布大小
        height=500,          # 高度可以固定，也可以尝试设置为 None 自适应
        xaxis_title='日期',
        yaxis_title='料量/kg')

    # 突出'均值'列，增加一列 alpha 或者 color
    highlight_unit = '均值'
    fig.update_traces(
        selector=lambda trace: trace.name == highlight_unit,
        line=dict(width=3, color='red'),    # 加粗
        marker=dict(size=8),                # 增大标记点
        opacity=1.0,                         # 提高不透明度
        hovertemplate=
        "<b>%{fullData.name}<br>"
        "<b>日龄</b>: %{customdata[0]} 天<br>"
        "<b>日期</b>: %{x:%Y-%m-%d}<br>"
        "<b>饲喂量</b>: %{y:.2f} kg"
        "<extra></extra>" # 鼠标悬停显示日龄和值
        
    )

    # 设置余下列
    fig.update_traces(
        selector=lambda trace: trace.name != highlight_unit,
        line=dict(width=1),
        opacity=0.6,
        hovertemplate=
        "<b>栏位</b>: %{fullData.name}<br>"
        "<b>日龄</b>: %{customdata[0]} 天<br>"
        "<b>日期</b>: %{x:%Y-%m-%d}<br>"
        "<b>饲喂量</b>: %{y:.2f} kg"
        "<extra></extra>" # 鼠标悬停显示日龄和值
    )

    # 设置 x 轴按天显示，日期标签旋转 90 度
    fig.update_xaxes(
        dtick="D1",        # 每天一个刻度
        tickangle=-90,      # 标签旋转90度
        tickformat="%m-%d", # 可选：只显示月-日
        tickfont=dict(size=10)  # 调小字体，例如10号
    )

    if save_img:
        fig.write_html(PATH_FIGURE_HTML / f'{unit} 28列单栏喂料量变化曲线.html', include_plotlyjs='cdn', full_html=True)
    return fig