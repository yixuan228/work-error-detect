import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager
import os
from datetime import datetime, timedelta
import matplotlib.dates as mdates

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'STHeiti', 'PingFang SC']
plt.rcParams['axes.unicode_minus'] = False

def read_accident_dates(filepath):
    """读取事故存栏文件，提取事故发生日期"""
    try:
        df = pd.read_excel(filepath)
        
        # 找到事故日期列
        date_col = '事故日期'
        if date_col not in df.columns:
            print(f"警告: 在文件 {filepath} 中未找到事故日期列")
            return []
        
        # 转换日期列
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        
        # 提取所有有效的事故日期（排除NaN和汇总行）
        accident_dates = df[date_col].dropna()
        
        # 只保留日期部分（去掉时间）
        accident_dates = accident_dates.dt.date.unique()
        
        # 转换回datetime以便在图表中使用
        accident_dates = [pd.Timestamp(date) for date in accident_dates]
        
        # 排序
        accident_dates = sorted(accident_dates)
        
        print(f"找到 {len(accident_dates)} 个事故发生日期")
        return accident_dates
        
    except Exception as e:
        print(f"读取事故文件 {filepath} 时出错: {e}")
        import traceback
        traceback.print_exc()
        return []

def read_transfer_and_sale_dates(filepath):
    """读取事故存栏文件Sheet1，提取转出和销售日期"""
    try:
        # 读取Sheet1工作表
        df = pd.read_excel(filepath, sheet_name='Sheet1')
        
        # Excel日期序列号从1899-12-30开始
        excel_epoch = datetime(1899, 12, 30)
        
        # 提取转出日期
        transfer_rows = df[df['转出'].notna() & (df['转出'] != 0)]
        transfer_dates = []
        for idx, row in transfer_rows.iterrows():
            date_num = row['日期']
            if pd.notna(date_num) and isinstance(date_num, (int, float)):
                date = excel_epoch + timedelta(days=int(date_num))
                transfer_dates.append(pd.Timestamp(date.date()))
        
        # 提取销售日期
        sale_rows = df[df['销售'].notna() & (df['销售'] != 0)]
        sale_dates = []
        for idx, row in sale_rows.iterrows():
            date_num = row['日期']
            if pd.notna(date_num) and isinstance(date_num, (int, float)):
                date = excel_epoch + timedelta(days=int(date_num))
                sale_dates.append(pd.Timestamp(date.date()))
        
        # 排序
        transfer_dates = sorted(transfer_dates)
        sale_dates = sorted(sale_dates)
        
        print(f"找到 {len(transfer_dates)} 个转出日期")
        print(f"找到 {len(sale_dates)} 个销售日期")
        return transfer_dates, sale_dates
        
    except Exception as e:
        print(f"读取转出/销售日期时出错: {e}")
        import traceback
        traceback.print_exc()
        return [], []

def read_treatment_dates(filepath):
    """读取事故存栏文件Sheet1，提取治疗日期和对应的数量"""
    try:
        # 读取Sheet1工作表
        df = pd.read_excel(filepath, sheet_name='Sheet1')
        
        # Excel日期序列号从1899-12-30开始
        excel_epoch = datetime(1899, 12, 30)
        
        # 提取治疗日期和数量（治疗列不为空且不为0的行）
        treatment_rows = df[df['治疗'].notna() & (df['治疗'] != 0)]
        treatment_dict = {}  # 使用字典存储日期和数量的对应关系
        
        for idx, row in treatment_rows.iterrows():
            date_num = row['日期']
            treatment_count = row['治疗']
            if pd.notna(date_num) and isinstance(date_num, (int, float)):
                date = excel_epoch + timedelta(days=int(date_num))
                date_key = pd.Timestamp(date.date())
                # 如果同一天有多个治疗记录，累加数量
                if date_key in treatment_dict:
                    treatment_dict[date_key] += float(treatment_count)
                else:
                    treatment_dict[date_key] = float(treatment_count)
        
        print(f"找到 {len(treatment_dict)} 个治疗日期")
        return treatment_dict
        
    except Exception as e:
        print(f"读取治疗日期时出错: {e}")
        import traceback
        traceback.print_exc()
        return {}

def read_and_process_by_pen(filepath):
    """读取Excel文件并按栏号处理数据"""
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

def create_visualization_by_pen(data, pen_number, filename, accident_dates, transfer_dates, sale_dates, treatment_dates):
    """为指定栏号创建带有事故/转出/销售/治疗日期标注的可视化图表"""
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
    
    # 创建图表
    fig, ax1 = plt.subplots(figsize=(18, 8))
    
    # 准备数据
    dates = pen_data['日期']
    feed_total = pen_data['每日采食总量']
    water_total = pen_data['每日喂水总量']
    water_feed_ratio = pen_data['水料比']
    
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
    ax1.xaxis.set_major_locator(mdates.DayLocator(interval=7))
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    
    # 设置次要刻度，显示所有日期（即使不显示标签）
    ax1.xaxis.set_minor_locator(mdates.DayLocator(interval=1))
    
    # 右Y轴：水料比（只显示实线）
    ax2 = ax1.twinx()
    color_ratio = 'blue'
    
    ax2.set_ylabel('水料比', fontsize=12, color=color_ratio)
    
    # 绘制水料比（实线）
    line5 = ax2.plot(dates, water_feed_ratio, color=color_ratio, linewidth=2, 
                     label='水料比', alpha=0.8)
    
    ax2.tick_params(axis='y', labelcolor=color_ratio)
    
    # 标注事故发生日期
    valid_accident_dates = []
    if accident_dates:
        # 获取数据日期范围
        date_min = dates.min()
        date_max = dates.max()
        
        # 筛选在数据日期范围内的事故日期
        valid_accident_dates = [d for d in accident_dates if date_min <= d <= date_max]
        
        if valid_accident_dates:
            # 为每个事故日期绘制垂直标注线
            for acc_date in valid_accident_dates:
                # 绘制红色垂直虚线标注事故发生日期
                ax1.axvline(x=acc_date, color='red', linestyle='--', linewidth=2, 
                           alpha=0.7)
            
            print(f"  标注了 {len(valid_accident_dates)} 个事故发生日期")
    
    # 标注转出日期
    valid_transfer_dates = []
    if transfer_dates:
        date_min = dates.min()
        date_max = dates.max()
        valid_transfer_dates = [d for d in transfer_dates if date_min <= d <= date_max]
        
        if valid_transfer_dates:
            for transfer_date in valid_transfer_dates:
                # 绘制绿色点划线标注转出日期
                ax1.axvline(x=transfer_date, color='green', linestyle='-.', linewidth=2, 
                           alpha=0.7, dashes=(5, 5))
            print(f"  标注了 {len(valid_transfer_dates)} 个转出日期")
    
    # 标注销售日期
    valid_sale_dates = []
    if sale_dates:
        date_min = dates.min()
        date_max = dates.max()
        valid_sale_dates = [d for d in sale_dates if date_min <= d <= date_max]
        
        if valid_sale_dates:
            for sale_date in valid_sale_dates:
                # 绘制蓝色点线标注销售日期
                ax1.axvline(x=sale_date, color='blue', linestyle=':', linewidth=2, 
                           alpha=0.7, dashes=(2, 2))
            print(f"  标注了 {len(valid_sale_dates)} 个销售日期")
    
    # 标注治疗日期
    valid_treatment_dates = {}
    if treatment_dates:
        date_min = dates.min()
        date_max = dates.max()
        # treatment_dates 是字典，需要提取日期键
        valid_treatment_dates = {d: treatment_dates[d] for d in treatment_dates 
                                if date_min <= d <= date_max}
        
        if valid_treatment_dates:
            # 使用ax2的Y轴范围（水料比）作为文本标注位置，避免与数据线重叠
            y2_min, y2_max = ax2.get_ylim()
            text_y_position = y2_max * 0.98  # 在右Y轴上方98%的位置显示文本
            
            for treatment_date, treatment_count in valid_treatment_dates.items():
                # 绘制紫色点划线标注治疗日期（使用不同的虚线样式）
                ax1.axvline(x=treatment_date, color='purple', linestyle='-.', linewidth=2, 
                           alpha=0.7, dashes=(8, 4, 2, 4))
                
                # 在治疗日期上方标注治疗数量（使用ax2以便在右Y轴上方显示）
                ax2.text(treatment_date, text_y_position, f'{int(treatment_count)}', 
                        color='purple', fontsize=8, fontweight='bold',
                        ha='center', va='bottom', rotation=0,
                        bbox=dict(boxstyle='round,pad=0.2', facecolor='white', 
                                 edgecolor='purple', alpha=0.9, linewidth=1.5))
            
            print(f"  标注了 {len(valid_treatment_dates)} 个治疗日期")
    
    # 合并图例，放在左上角避免与折线重叠
    lines = line1 + line2 + line3 + line4 + line5
    labels = [l.get_label() for l in lines]
    
    # 添加各种日期标注到图例
    from matplotlib.lines import Line2D
    if accident_dates and valid_accident_dates:
        accident_line = Line2D([0], [0], color='red', linestyle='--', linewidth=2, 
                              label='事故发生日期')
        lines.append(accident_line)
        labels.append('事故发生日期')
    
    if transfer_dates and valid_transfer_dates:
        transfer_line = Line2D([0], [0], color='green', linestyle='-.', linewidth=2, 
                              label='转出日期', dashes=(5, 5))
        lines.append(transfer_line)
        labels.append('转出日期')
    
    if sale_dates and valid_sale_dates:
        sale_line = Line2D([0], [0], color='blue', linestyle=':', linewidth=2, 
                          label='销售日期', dashes=(2, 2))
        lines.append(sale_line)
        labels.append('销售日期')
    
    if treatment_dates and valid_treatment_dates:
        treatment_line = Line2D([0], [0], color='purple', linestyle='-.', linewidth=2, 
                              label='治疗日期（含数量）', dashes=(8, 4, 2, 4))
        lines.append(treatment_line)
        labels.append('治疗日期（含数量）')
    
    # 将图例放在左上角，确保不遮挡数据
    legend = ax1.legend(lines, labels, loc='upper left', fontsize=10, 
                       framealpha=0.95, fancybox=True, shadow=True)
    
    # 设置标题
    unit_name = filename.replace('.xlsx', '').replace(' (1)', '')
    plt.title(f'{unit_name} - 第{pen_number}栏 - 饲喂数据可视化（标注事故/转出/销售/治疗日期）', fontsize=14, fontweight='bold', pad=20)
    
    # 格式化x轴日期，垂直显示（旋转90度）避免重叠
    plt.xticks(rotation=90, ha='center', va='top')
    
    # 调整布局，为日期标签留出更多空间
    plt.subplots_adjust(bottom=0.25)
    
    # 保存图表到新文件夹
    output_dir = '4栋2单元-含事件'
    os.makedirs(output_dir, exist_ok=True)
    output_filename = f'{unit_name}_第{pen_number}栏_可视化图表.png'
    output_path = os.path.join(output_dir, output_filename)
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    print(f"已保存图表: {output_path}")
    
    plt.close()

# 读取事故日期、转出日期、销售日期和治疗日期
accident_file = '4栋2单元事故存栏 全阶段.xls'
print(f"正在读取事故文件: {accident_file}")
accident_dates = read_accident_dates(accident_file)
transfer_dates, sale_dates = read_transfer_and_sale_dates(accident_file)
treatment_dates = read_treatment_dates(accident_file)

# 处理育肥4-2单元文件
file = '育肥4-2单元饲喂记录-2025-08-25--2025-12-16.xlsx'

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
            create_visualization_by_pen(data, pen_num, file, accident_dates, transfer_dates, sale_dates, treatment_dates)
        
        print(f"\n所有图表已生成完成！共生成 {len(target_pens)} 个图表，保存在 '4栋2单元-含事件' 文件夹中。")
    else:
        print("无法读取数据")
else:
    print(f"文件不存在: {file}")

