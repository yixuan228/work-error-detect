from data_etl.utils.process_single_file import process_single_file

from data_etl.config.path import PATH_FEED_ORI

import pandas as pd
from pathlib import Path

from tqdm import tqdm
# 处理原始数据，更新数据库
for file in tqdm(PATH_FEED_ORI.glob('*.xlsx'), desc='Processing'):
    process_single_file(file.stem)

# Terminal: python -m feed_analysis.scripts.xlsx_2parquet