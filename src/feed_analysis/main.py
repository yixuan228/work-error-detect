from feed_analysis.feed_pipeline.utils.process_single_file import process_single_file
from feed_analysis.feed_pipeline.utils.age import recalibrate_age
from feed_analysis.config.path import PATH_FEED_ORI, PATH_FEED_PROCESSED

import pandas as pd
from pathlib import Path

def main():
    # 处理原始数据，更新数据库
    for file in PATH_FEED_ORI.glob('*.xlsx'):
        process_single_file(file.stem)

    # 矫正日期
    file_count = len([f for f in PATH_FEED_PROCESSED.iterdir() if f.is_file()])
    dates = ['2025-8-28'] * file_count                                              # 需要手动调整
    ages = [32] * file_count                                                        # 需要手动调整
    for file in PATH_FEED_PROCESSED.glob('*.parquet'):
        recalibrate_age(file, reference_date=dates[file_count - 1], reference_age=ages[file_count - 1])

if __name__ == '__main__':
    main()  

    # Terminal: python -m feed_analysis.main