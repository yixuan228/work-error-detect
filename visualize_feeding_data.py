import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager
import os
from datetime import datetime
import matplotlib.dates as mdates

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'STHeiti', 'PingFang SC']
plt.rcParams['axes.unicode_minus'] = False

# 文件列表
files = [
    '育肥3-1单元饲喂记录-2025-08-28--2025-12-10.xlsx',
    '育肥3-2单元饲喂记录-2025-08-28--2025-12-10 (1).xlsx',
    '育肥3-3单元饲喂记录-2025-08-28--2025-12-10.xlsx',
    '育肥3-4单元饲喂记录-2025-08-28--2025-12-10.xlsx',
    '育肥4-2单元饲喂记录-2025-08-25--2025-12-16.xlsx'
]

def read_and_process_file(filepath):
    """读取Excel文件并处理数据"""
    try:
        # 读取Excel文件，跳过前3行，使用第3行（索引3）作为列名
        df = pd.read_excel(filepath, header=3)
        
        # 清理列名，去除前后空格
        df.columns = df.columns.str.strip()
        
        # 确保日期列存在
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
        
        # 转换日期列
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        
        # 找到每日采食总量和每日喂水总量列
        feed_col = None
        water_col = None
        
        for col in df.columns:
            if '每日采食总量' in str(col) or ('采食' in str(col) and '总量' in str(col)):
                feed_col = col
            if '每日喂水总量' in str(col) or ('喂水' in str(col) and '总量' in str(col)):
                water_col = col
        
        if feed_col is None or water_col is None:
            print(f"警告: 在文件 {filepath} 中未找到所需列")
            print(f"可用列: {df.columns.tolist()}")
            return None
        
        # 按日期分组，取每日的唯一值（因为每天的数据中，每日总量应该是相同的）
        daily_data = df.groupby(date_col).agg({
            feed_col: 'first',
            water_col: 'first'
        }).reset_index()
        
        # 重命名列
        daily_data.columns = ['日期', '每日采食总量', '每日喂水总量']
        
        # 删除缺失值
        daily_data = daily_data.dropna()
        
        # 按日期排序
        daily_data = daily_data.sort_values('日期')
        
        # 计算水料比
        daily_data['水料比'] = daily_data['每日喂水总量'] / daily_data['每日采食总量']
        
        return daily_data
        
    except Exception as e:
        print(f"读取文件 {filepath} 时出错: {e}")
        import traceback
        traceback.print_exc()
        return None

def create_visualization(data, filename):
    """创建可视化图表"""
    if data is None or len(data) == 0:
        print(f"无法为 {filename} 创建图表：数据为空")
        return
    
    # 创建图表，增大图形尺寸以容纳更多日期标签
    fig, ax1 = plt.subplots(figsize=(18, 8))
    
    # 准备数据
    dates = data['日期']
    feed_total = data['每日采食总量']
    water_total = data['每日喂水总量']
    water_feed_ratio = data['水料比']
    
    # 计算线性趋势
    x_numeric = np.arange(len(dates))
    
    # 每日采食总量的线性趋势
    feed_coef = np.polyfit(x_numeric, feed_total, 1)
    feed_trend = np.polyval(feed_coef, x_numeric)
    
    # 每日喂水总量的线性趋势
    water_coef = np.polyfit(x_numeric, water_total, 1)
    water_trend = np.polyval(water_coef, x_numeric)
    
    # 左Y轴：每日采食总量和每日喂水总量
    color_feed = 'gray'
    color_water = '#FFA500'  # 橙色/黄色
    
    ax1.set_xlabel('日期', fontsize=12)
    ax1.set_ylabel('每日采食总量(Kg) / 每日喂水总量(L)', fontsize=12)
    
    # 绘制每日采食总量（实线）
    line1 = ax1.plot(dates, feed_total, color=color_feed, linewidth=2, 
                     label='每日采食总量(Kg)', alpha=0.8)
    
    # 绘制每日采食总量的线性趋势（虚线）
    line2 = ax1.plot(dates, feed_trend, color=color_feed, linewidth=2, 
                     linestyle='--', label='线性(每日采食总量(Kg))', alpha=0.6)
    
    # 绘制每日喂水总量（实线）
    line3 = ax1.plot(dates, water_total, color=color_water, linewidth=2, 
                     label='每日喂水总量(L)', alpha=0.8)
    
    # 绘制每日喂水总量的线性趋势（虚线）
    line4 = ax1.plot(dates, water_trend, color=color_water, linewidth=2, 
                     linestyle='--', label='线性(每日喂水总量(L))', alpha=0.6)
    
    ax1.tick_params(axis='y')
    ax1.grid(True, alpha=0.3)
    
    # 设置x轴日期格式，按每周显示（每7天显示一次）
    # 使用DayLocator每7天显示一次，确保不会太密集
    ax1.xaxis.set_major_locator(mdates.DayLocator(interval=7))
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    
    # 设置次要刻度，显示所有日期（即使不显示标签）
    ax1.xaxis.set_minor_locator(mdates.DayLocator(interval=1))
    
    # 右Y轴：水料比（只显示实线，不显示虚线）
    ax2 = ax1.twinx()
    color_ratio = 'blue'
    
    ax2.set_ylabel('水料比', fontsize=12, color=color_ratio)
    
    # 绘制水料比（实线）
    line5 = ax2.plot(dates, water_feed_ratio, color=color_ratio, linewidth=2, 
                     label='水料比', alpha=0.8)
    
    ax2.tick_params(axis='y', labelcolor=color_ratio)
    
    # 合并图例，放在左上角避免与折线重叠
    lines = line1 + line2 + line3 + line4 + line5
    labels = [l.get_label() for l in lines]
    
    # 将图例放在左上角，确保不遮挡数据
    legend = ax1.legend(lines, labels, loc='upper left', fontsize=10, 
                       framealpha=0.95, fancybox=True, shadow=True)
    
    # 设置标题
    unit_name = filename.replace('.xlsx', '').replace(' (1)', '')
    plt.title(f'{unit_name} - 饲喂数据可视化', fontsize=14, fontweight='bold', pad=20)
    
    # 格式化x轴日期，垂直显示（旋转90度）避免重叠
    plt.xticks(rotation=90, ha='center', va='top')
    
    # 调整布局，为日期标签留出更多空间
    # 增加底部边距以容纳垂直的日期标签
    plt.subplots_adjust(bottom=0.25)
    
    # 保存图表
    output_filename = f'{unit_name}_可视化图表.png'
    plt.savefig(output_filename, dpi=300, bbox_inches='tight')
    print(f"已保存图表: {output_filename} (日期间隔: 每周显示)")
    
    plt.close()

# 处理所有文件
for file in files:
    if os.path.exists(file):
        print(f"\n正在处理: {file}")
        data = read_and_process_file(file)
        if data is not None:
            print(f"成功读取 {len(data)} 条记录")
            create_visualization(data, file)
    else:
        print(f"文件不存在: {file}")

print("\n所有图表已生成完成！")

