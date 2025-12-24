import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager
import os
from datetime import datetime, timedelta
import matplotlib.dates as mdates
import re

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

def read_henan_standard(filepath):
    """读取河南的饲喂标准数据"""
    try:
        # 读取Excel文件
        df = pd.read_excel(filepath, header=None)
        
        # 找到河南数据开始的行（包含"河南"的行）
        henan_start_idx = None
        for idx, row in df.iterrows():
            if pd.notna(row[1]) and '河南' in str(row[1]):
                henan_start_idx = idx
                break
        
        if henan_start_idx is None:
            print("警告: 未找到河南数据")
            return None
        
        # 读取河南的数据（从标题行开始，标题行是henan_start_idx）
        # 数据从henan_start_idx+1开始
        data_rows = df.iloc[henan_start_idx+1:].copy()
        
        # 解析阶段和采食量
        standards = []
        for idx, row in data_rows.iterrows():
            stage_str = str(row[3]) if pd.notna(row[3]) else ''
            daily_feed = row[5] if pd.notna(row[5]) else None
            
            # 跳过合计行和空行
            if '合计' in stage_str or daily_feed is None or pd.isna(daily_feed):
                continue
            
            # 解析阶段日龄范围，例如 "0-1月龄（26-30）"
            age_match = re.search(r'（(\d+)-(\d+)）', stage_str)
            if age_match:
                age_start = int(age_match.group(1))
                age_end = int(age_match.group(2))
                try:
                    daily_feed_value = float(daily_feed)
                    standards.append({
                        'age_start': age_start,
                        'age_end': age_end,
                        'daily_feed_per_head': daily_feed_value
                    })
                except (ValueError, TypeError):
                    continue
        
        # 按日龄排序
        standards.sort(key=lambda x: x['age_start'])
        
        return standards
        
    except Exception as e:
        print(f"读取标准文件 {filepath} 时出错: {e}")
        import traceback
        traceback.print_exc()
        return None

def get_standard_feed_by_age(standards, age):
    """根据日龄获取标准头均日采食量"""
    for std in standards:
        if std['age_start'] <= age <= std['age_end']:
            return std['daily_feed_per_head']
    # 如果超出范围，使用边界值
    if standards:
        if age < standards[0]['age_start']:
            return standards[0]['daily_feed_per_head']
        elif age > standards[-1]['age_end']:
            return standards[-1]['daily_feed_per_head']
    return None

def read_and_process_file(filepath):
    """读取Excel文件并处理数据，同时获取总猪只数"""
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
            return None, None
        
        # 转换日期列
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        
        # 找到每日采食总量和每日喂水总量列
        feed_col = None
        water_col = None
        head_count_col = None
        pen_col = None
        
        for col in df.columns:
            col_str = str(col)
            if '每日采食总量' in col_str or ('采食' in col_str and '总量' in col_str):
                feed_col = col
            if '每日喂水总量' in col_str or ('喂水' in col_str and '总量' in col_str):
                water_col = col
            if '猪只头数' in col_str or '头数' in col_str:
                head_count_col = col
            if '栏号' in col_str or '栏' in col_str:
                pen_col = col
        
        if feed_col is None or water_col is None:
            print(f"警告: 在文件 {filepath} 中未找到所需列")
            print(f"可用列: {df.columns.tolist()}")
            return None, None
        
        # 计算总猪只数（栏数 × 每栏头数）
        total_head_count = None
        if head_count_col and pen_col:
            # 获取每栏的猪只头数（应该是一致的）
            pen_headcount = df.groupby(pen_col)[head_count_col].first()
            total_pens = len(pen_headcount)
            head_count_per_pen = pen_headcount.iloc[0] if len(pen_headcount) > 0 else None
            if head_count_per_pen:
                total_head_count = total_pens * head_count_per_pen
                print(f"总栏数: {total_pens}, 每栏猪只头数: {head_count_per_pen}, 总猪只数: {total_head_count}")
        
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
        
        return daily_data, total_head_count
        
    except Exception as e:
        print(f"读取文件 {filepath} 时出错: {e}")
        import traceback
        traceback.print_exc()
        return None, None

def create_visualization_with_accidents(data, accident_dates, transfer_dates, sale_dates, treatment_dates, filename, standards, total_head_count, start_age=25, end_age=114):
    """创建带有事故日期、转出日期、销售日期、治疗日期标注和标准曲线的可视化图表"""
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
    
    # 计算标准曲线
    standard_feed = None
    standard_water = None
    standard_dates = None
    
    if standards and total_head_count:
        # 计算每个日期对应的日龄
        # 假设第一个日期对应转舍日龄（start_age）
        first_date = dates.iloc[0]
        date_to_age = {}
        
        for i, date in enumerate(dates):
            # 计算从第一个日期到当前日期的天数
            days_from_start = (date - first_date).days
            # 计算当前日龄
            current_age = start_age + days_from_start
            
            # 如果超出范围，使用边界值
            if current_age < start_age:
                current_age = start_age
            elif current_age > end_age:
                current_age = end_age
            
            date_to_age[date] = current_age
        
        # 计算标准采食量和饮水量
        standard_feed_list = []
        standard_water_list = []
        standard_dates_list = []
        
        for date in dates:
            age = date_to_age[date]
            daily_feed_per_head = get_standard_feed_by_age(standards, age)
            if daily_feed_per_head:
                # 计算总的标准采食量（头均日采食量 × 总猪只数）
                standard_feed_total = daily_feed_per_head * total_head_count
                # 计算总的标准饮水量（头均日采食量 × 3 × 总猪只数）
                standard_water_total = daily_feed_per_head * 3 * total_head_count
                
                standard_feed_list.append(standard_feed_total)
                standard_water_list.append(standard_water_total)
                standard_dates_list.append(date)
        
        if standard_feed_list:
            standard_feed = np.array(standard_feed_list)
            standard_water = np.array(standard_water_list)
            standard_dates = pd.Series(standard_dates_list)
    
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
    
    # 绘制标准曲线
    color_standard_feed = 'green'
    color_standard_water = 'darkgreen'
    
    lines = line1 + line2 + line3 + line4 + line5
    if standard_feed is not None:
        line6 = ax1.plot(standard_dates, standard_feed, color=color_standard_feed, 
                        linewidth=2.5, linestyle=':', label='标准每日采食总量(Kg)', alpha=0.9)
        lines = lines + line6
    
    if standard_water is not None:
        line7 = ax1.plot(standard_dates, standard_water, color=color_standard_water, 
                        linewidth=2.5, linestyle=':', label='标准每日喂水总量(L)', alpha=0.9)
        lines = lines + line7
    
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
                # 绘制红色垂直虚线标注事故发生日期（不添加标记图案）
                ax1.axvline(x=acc_date, color='red', linestyle='--', linewidth=2, 
                           alpha=0.7, label='事故发生日期' if acc_date == valid_accident_dates[0] else '')
            
            print(f"标注了 {len(valid_accident_dates)} 个事故发生日期")
    
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
            print(f"标注了 {len(valid_transfer_dates)} 个转出日期")
    
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
            print(f"标注了 {len(valid_sale_dates)} 个销售日期")
    
    # 标注治疗日期
    valid_treatment_dates = []
    if treatment_dates:
        date_min = dates.min()
        date_max = dates.max()
        # treatment_dates 现在是字典，需要提取日期键
        valid_treatment_dates = {d: treatment_dates[d] for d in treatment_dates 
                                 if date_min <= d <= date_max}
        
        if valid_treatment_dates:
            # 获取Y轴范围，用于定位文本标注位置
            y_min, y_max = ax1.get_ylim()
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
            
            print(f"标注了 {len(valid_treatment_dates)} 个治疗日期")
    
    # 合并图例，放在左上角避免与折线重叠
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
    plt.title(f'{unit_name} - 饲喂数据可视化（含标准曲线及事件标注）', fontsize=14, fontweight='bold', pad=20)
    
    # 格式化x轴日期，垂直显示（旋转90度）避免重叠
    plt.xticks(rotation=90, ha='center', va='top')
    
    # 调整布局，为日期标签留出更多空间
    plt.subplots_adjust(bottom=0.25)
    
    # 保存图表
    output_filename = f'{unit_name}_可视化图表_含标准曲线及事件标注.png'
    plt.savefig(output_filename, dpi=300, bbox_inches='tight')
    print(f"已保存图表: {output_filename} (总猪只数: {total_head_count if total_head_count else '未知'})")
    
    plt.close()

# 读取河南饲喂标准
standard_file = '商品猪各阶段料量2025.11.13.xlsx'
print(f"正在读取标准文件: {standard_file}")
standards = read_henan_standard(standard_file)
if standards:
    print(f"成功读取 {len(standards)} 个阶段的河南饲喂标准")

# 读取事故日期、转出日期、销售日期和治疗日期
accident_file = '4栋2单元事故存栏 全阶段.xls'
print(f"\n正在读取事故文件: {accident_file}")
accident_dates = read_accident_dates(accident_file)
transfer_dates, sale_dates = read_transfer_and_sale_dates(accident_file)
treatment_dates = read_treatment_dates(accident_file)

# 处理育肥4-2单元文件
feeding_file = '育肥4-2单元饲喂记录-2025-08-25--2025-12-16.xlsx'

if os.path.exists(feeding_file):
    print(f"\n正在处理: {feeding_file}")
    data, total_head_count = read_and_process_file(feeding_file)
    if data is not None:
        print(f"成功读取 {len(data)} 条记录")
        create_visualization_with_accidents(data, accident_dates, transfer_dates, sale_dates, 
                                           treatment_dates, feeding_file, standards, total_head_count, 
                                           start_age=25, end_age=114)
        print("\n图表已生成完成！")
    else:
        print("无法读取数据")
else:
    print(f"文件不存在: {feeding_file}")

