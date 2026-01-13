from matplotlib import font_manager
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
CN_font = font_manager.FontProperties(fname='C:/Windows/Fonts/simhei.ttf')

def plt_feed_line_build(df, unit:str, column, if_partial=False, start_date=None, end_date=None, save_img=False):
    
    # df = pd.read_excel(f'data/feed_record/{file_name}.xlsx', index_col=0, header=3).loc[::-1, ['转舍前日龄', '每日采食总量(Kg)', '每日喂水总量(L)']].drop_duplicates()
    df.index = pd.to_datetime(df.index)

    if if_partial:
        df = df.loc[df['Date'].between(pd.to_datetime(start_date).date(), pd.to_datetime(end_date).date()), :]

    plt.figure(figsize=(12, 5))
    plt.plot(df['Date'], df['food_total_kg'], label='日摄入量', color='red')
    plt.xlabel('日期', fontproperties=CN_font)
    plt.ylabel('饲喂量/kg', fontproperties=CN_font)
    plt.title(f'{unit}料/水总量', fontproperties=CN_font)

    # 图细节
    ax = plt.gca()
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    ax.grid(axis='x', linestyle='--', color='gray', alpha=0.4)  # 只显示 x 轴网格线

    plt.xticks(rotation=90, fontproperties=CN_font, fontsize=7)
    plt.legend()
    plt.tight_layout()
    plt.savefig(f'figure/{unit}-饲料总量变化.png', dpi=300, bbox_inches='tight') if save_img else None
    plt.show()