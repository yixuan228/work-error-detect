# 所有路径配置

from pathlib import Path

# 项目根目录
PATH_ROOT = Path(__file__).parent.parent.parent.parent

# 日志目录
PATH_LOG = PATH_ROOT / 'logs'

# 数据目录
PATH_DATA = PATH_ROOT / 'data'

# 喂食数据目录
PATH_FEED_ORI = PATH_DATA / 'feed_record_ori'
PATH_FEED_PROCESSED  = PATH_DATA / 'feed_record_processed'