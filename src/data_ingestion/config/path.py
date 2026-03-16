
from pathlib import Path

# 项目根目录
PATH_ROOT = Path(__file__).parent.parent.parent.parent

# 数据目录
PATH_DATA = PATH_ROOT / 'data'

# 喂食数据目录
PATH_FEED_ORI = PATH_DATA / 'feed_record_ori'                   # 原始数据位置
PATH_FEED_PROCESSED  = PATH_DATA / 'feed_record_processed'      # 处理后parquet位置

# src目录
PATH_SRC_INGEST = PATH_ROOT / 'src' / 'data_ingestion'  # src模块的路径
