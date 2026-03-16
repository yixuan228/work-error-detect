from data_etl.utils.age import recalibrate_age
from data_etl.config.path import PATH_FEED_PROCESSED

# 矫正日期
# file_count = len([f for f in PATH_FEED_PROCESSED.iterdir() if f.is_file()])
# dates = ['2025-8-28'] * file_count                                              # 需要手动调整
# ages = [32] * file_count                                                        # 需要手动调整
# for file in PATH_FEED_PROCESSED.glob('*.parquet'):
#     recalibrate_age(file, reference_date=dates[file_count - 1], reference_age=ages[file_count - 1])

recalibrate_age(PATH_FEED_PROCESSED / '育肥5-3_build.parquet', reference_date='2025-10-15', reference_age=24)