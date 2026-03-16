import pandas as pd
import matplotlib
matplotlib.use('Agg')  # 使用非交互式后端
import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager
from matplotlib.ticker import FuncFormatter
import os
from datetime import datetime
import matplotlib.dates as mdates

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'STHeiti', 'PingFang SC']
plt.rcParams['axes.unicode_minus'] = False

def read_and_process_by_pen(filepath):
    """读取Excel文件并按栏号处理数据，计算每日的采食总量和喂水总量"""
    try:
        # 读取Excel文件，跳过前3行，使用第3行（索引3）作为列名
        df = pd.read_excel(filepath, header=3)
        
        # 清理列名，去除前后空格
        df.columns = df.columns.str.strip()
        
        # 找到日期列
        date_col = None
        for col in df.columns:
            col_str = str(col)
            if '日期' in col_str or 'date' in col_str.lower() or '饲喂日期' in col_str:
                date_col = col
                break
        
        if date_col is None:
            print(f"警告: 在文件 {filepath} 中未找到日期列")
            print(f"可用列: {df.columns.tolist()}")
            return None
        
        # 找到栏号列
        pen_col = None
        for col in df.columns:
            col_str = str(col)
            if '栏号' in col_str or '栏' in col_str:
                pen_col = col
                break
        
        if pen_col is None:
            print(f"警告: 在文件 {filepath} 中未找到栏号列")
            print(f"可用列: {df.columns.tolist()}")
            return None
        
        # 找到单栏采食量和单栏喂水量列
        feed_col = None
        water_col = None
        
        for col in df.columns:
            col_str = str(col)
            if '单栏' in col_str and ('采食' in col_str or '喂料' in col_str or '料量' in col_str):
                feed_col = col
            if '单栏' in col_str and ('喂水' in col_str or '饮水' in col_str or '水量' in col_str):
                water_col = col
        
        if feed_col is None or water_col is None:
            print(f"警告: 在文件 {filepath} 中未找到单栏采食量或单栏喂水量列")
            print(f"可用列: {df.columns.tolist()}")
            return None
        
        print(f"找到的列: 日期={date_col}, 栏号={pen_col}, 喂料量={feed_col}, 喂水量={water_col}")
        
        # 转换日期列
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        
        # 删除缺失值
        df = df.dropna(subset=[date_col, pen_col, feed_col, water_col])
        
        # 按栏号和日期分组，计算每日的采食总量和喂水总量
        daily_data = df.groupby([pen_col, date_col]).agg({
            feed_col: 'sum',  # 如果同一天同一栏有多条记录，求和
            water_col: 'sum'
        }).reset_index()
        
        # 重命名列
        daily_data.columns = ['栏号', '日期', '每日采食总量', '每日喂水总量']
        
        # 按栏号和日期排序
        daily_data = daily_data.sort_values(['栏号', '日期'])
        
        # 计算水料比
        daily_data['水料比'] = daily_data['每日喂水总量'] / daily_data['每日采食总量']
        
        return daily_data
        
    except Exception as e:
        print(f"读取文件 {filepath} 时出错: {e}")
        import traceback
        traceback.print_exc()
        return None

def format_date_label(date):
    """格式化日期标签：2025-12-15只显示12-15，其他显示完整日期"""
    if date.year == 2025 and date.month == 12 and date.day == 15:
        return f"{date.month}-{date.day}"
    else:
        return date.strftime('%Y-%m-%d')

def create_visualization_by_pen(data, pen_number, filename, output_dir):
    """为指定栏号创建可视化图表"""
    if data is None or len(data) == 0:
        print(f"无法为栏号 {pen_number} 创建图表：数据为空")
        return
    
    # 筛选该栏号的数据
    pen_data = data[data['栏号'] == pen_number].copy()
    
    if len(pen_data) == 0:
        print(f"栏号 {pen_number} 没有数据")
        return
    
    # 按日期排序
    pen_data = pen_data.sort_values('日期')
    
    # 创建图表，增大图形尺寸以容纳更多日期标签
    fig, ax1 = plt.subplots(figsize=(20, 10))
    
    # 准备数据
    dates = pen_data['日期']
    feed_total = pen_data['每日采食总量']
    water_total = pen_data['每日喂水总量']
    water_feed_ratio = pen_data['水料比']
    
    # 左Y轴：每日采食总量和每日喂水总量
    color_feed = 'gray'
    color_water = '#FFA500'  # 橙色/黄色
    
    ax1.set_xlabel('日期', fontsize=14, fontweight='bold')
    ax1.set_ylabel('每日采食总量(Kg) / 每日喂水总量(L)', fontsize=14, fontweight='bold')
    
    # 绘制每日采食总量（灰色实线）
    line1 = ax1.plot(dates, feed_total, color=color_feed, linewidth=2.5, 
                     label='每日采食总量(Kg)', alpha=0.9)
    
    # 绘制每日喂水总量（橙色/黄色实线）
    line3 = ax1.plot(dates, water_total, color=color_water, linewidth=2.5, 
                     label='每日喂水总量(L)', alpha=0.9)
    
    ax1.tick_params(axis='y', labelsize=12)
    ax1.grid(True, alpha=0.3, linestyle='-', linewidth=0.5)
    
    # 设置x轴日期格式，显示所有日期，竖放
    # 使用DayLocator显示每天
    ax1.xaxis.set_major_locator(mdates.DayLocator(interval=1))
    
    # 自定义日期格式化函数
    def date_formatter(x, pos=None):
        date = mdates.num2date(x)
        return format_date_label(date)
    
    ax1.xaxis.set_major_formatter(FuncFormatter(date_formatter))
    
    # 获取所有日期刻度并设置旋转和对齐
    ax1.tick_params(axis='x', labelsize=8)  # 减小字体大小
    
    # 右Y轴：水料比
    ax2 = ax1.twinx()
    color_ratio = 'blue'
    
    ax2.set_ylabel('水料比（每日喂水量/每日采食量）', fontsize=14, fontweight='bold', color=color_ratio)
    
    # 绘制水料比（蓝色实线）
    line5 = ax2.plot(dates, water_feed_ratio, color=color_ratio, linewidth=2.5, 
                     label='水料比', alpha=0.9)
    
    ax2.tick_params(axis='y', labelcolor=color_ratio, labelsize=12)
    
    # 合并图例
    lines = line1 + line3 + line5
    labels = [l.get_label() for l in lines]
    
    # 将图例放在左上角
    legend = ax1.legend(lines, labels, loc='upper left', fontsize=11, 
                       framealpha=0.95, fancybox=True, shadow=True)
    
    # 设置标题
    unit_name = filename.replace('.xlsx', '').replace(' (1)', '')
    plt.title(f'{unit_name} - 第{pen_number}栏 - 每日采食总量、喂水总量及水料比', 
              fontsize=16, fontweight='bold', pad=20)
    
    # 格式化x轴日期，垂直显示（旋转90度）避免重叠
    # 设置所有日期标签竖放，使用右对齐和顶部对齐
    for label in ax1.get_xticklabels():
        label.set_rotation(90)
        label.set_ha('right')
        label.set_va('top')
        label.set_fontsize(8)  # 使用较小的字体
    
    # 调整布局，为日期标签留出更多空间（根据日期数量动态调整底部边距）
    num_dates = len(dates)
    if num_dates > 80:
        bottom_margin = 0.35  # 日期很多时，需要更多底部空间
    elif num_dates > 50:
        bottom_margin = 0.3
    else:
        bottom_margin = 0.25
    
    plt.subplots_adjust(bottom=bottom_margin)
    
    # 保存图表到新文件夹
    os.makedirs(output_dir, exist_ok=True)
    output_filename = f'{unit_name}_第{pen_number}栏_可视化图表.png'
    output_path = os.path.join(output_dir, output_filename)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    print(f"已保存图表: {output_path}")
    
    plt.close()

# 主程序
if __name__ == "__main__":
    # 处理育肥4-2单元文件
    file = '育肥4-2单元饲喂记录-2025-08-25--2025-12-16.xlsx'
    output_dir = '4栋2单元-每日采食喂水水料比图表'
    
    if os.path.exists(file):
        print(f"\n正在处理: {file}")
        data = read_and_process_by_pen(file)
        
        if data is not None:
            # 获取所有栏号（1-28栏）
            all_pens = sorted(data['栏号'].unique())
            # 确保只生成1-28栏
            target_pens = [p for p in all_pens if 1 <= p <= 28]
            
            print(f"\n将为以下 {len(target_pens)} 个栏号生成图表: {target_pens}")
            
            # 为所有栏号生成图表
            for pen_num in target_pens:
                print(f"\n正在生成第 {pen_num} 栏的图表...")
                create_visualization_by_pen(data, pen_num, file, output_dir)
            
            print(f"\n所有图表已生成完成！共生成 {len(target_pens)} 个图表，保存在 '{output_dir}' 文件夹中。")
        else:
            print("无法读取数据")
    else:
        print(f"文件不存在: {file}")

