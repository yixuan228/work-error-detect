from matplotlib import font_manager
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
CN_font = font_manager.FontProperties(fname='C:/Windows/Fonts/simhei.ttf')

# 静态 单个 单元饲料总量变化折线图
def line_static_feed_build(df, unit:str, if_partial=False, start_date=None, end_date=None, save_img=False):
    """"
    静态单个单元层饲料总量变化折线图绘制
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
    plt.savefig(f'figure/{unit}-饲料总量变化.png', dpi=300, bbox_inches='tight') if save_img else None
    plt.show()


# 动态 多单元 饲料总量变化折线图
import plotly.express as px
def line_dynamic_feed_build(df_list:list, name_list:list, if_partial=False, start_date=None, end_date=None, save_img=False):
    """
    动态多单元层饲料总量（包含平均值）变化折线图绘制
    """
    ## ------ 读取数据部分 ------
    # 设置第一个表格（先根据日期和采食量去重，再设置日期为索引）
    df = df_list[0].loc[:, ['Date', 'food_total_kg']]
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

    if if_partial:
        df = df[(df.index >= pd.to_datetime(start_date)) & (df.index <= pd.to_datetime(end_date))]

    ## ------ 数据转换部分 宽数据转长数据 ------
    df_long = df.reset_index().melt(id_vars='Date', var_name='Unit', value_name='Food Intake (Kg)')

    ## ------ 绘图部分 ------
    # 交互图像

    # 绘制交互曲线
    fig = px.line(df_long, 
                x='Date', y='Food Intake (Kg)', 
                color='Unit', 
                title=f'各单元总料量变化趋势曲线', )

    # 突出显示均值曲线
    highlight_unit = '均值'
    fig.update_traces(
        selector=lambda trace: trace.name == highlight_unit,
        line=dict(width=3, color='red'),    # 加粗
        marker=dict(size=8),                # 增大标记点
        opacity=0.6                         # 提高不透明度
    )

    # 更新其余曲线
    fig.update_traces(
        selector=lambda trace: trace.name != highlight_unit,
        line=dict(width=1),
        opacity=0.6
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
        fig.write_html(f'figure/三栋1-4单元总料量变化趋势曲线.html', include_plotlyjs='cdn', full_html=True)

    return fig

# 动态 单单元
import plotly.express as px
def line_dynamic_feed_column(df, unit:str, if_partial=False, start_date=None, end_date=None, save_img=False):
    """
    动态单个单元 栏位级饲料总量（包含平均值）变化折线图绘制
    """ 
    if if_partial:
        df = df.loc[df['Date'].between(pd.to_datetime(start_date).date(), pd.to_datetime(end_date).date()), :]

    # 转宽表矩阵，添加Total
    df = df[['Date', 'col_num', 'food_col_kg']]
    df_wide = df.pivot(index='Date', columns='col_num', values='food_col_kg')
    df_wide['均值'] = df_wide.mean(axis=1)
    df_long = df_wide.reset_index().melt(id_vars='Date', var_name='Column', value_name='value')

    fig = px.line(df_long, x='Date', y='value', color='Column', title=f'{unit} 28列单栏喂料量变化曲线')
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
        opacity=1.0                         # 提高不透明度
    )

    # 设置余下列
    fig.update_traces(
        selector=lambda trace: trace.name != highlight_unit,
        line=dict(width=1),
        opacity=0.6
    )

    # 设置 x 轴按天显示，日期标签旋转 90 度
    fig.update_xaxes(
        dtick="D1",        # 每天一个刻度
        tickangle=-90,      # 标签旋转90度
        tickformat="%m-%d", # 可选：只显示月-日
        tickfont=dict(size=10)  # 调小字体，例如10号
    )

    if save_img:
        fig.write_html(f'figure/{unit} 28列单栏喂料量变化曲线.html', include_plotlyjs='cdn', full_html=True)
    return fig