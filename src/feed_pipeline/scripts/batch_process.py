from feed_pipeline.meta.feed_file_meta import FeedFileMeta
from feed_pipeline.io.reader import read_feed_excel
from feed_pipeline.io.writer import save_incremental_parquet
from feed_pipeline.processing.age import calc_relative_age
from feed_pipeline.processing.aggregation import split_granularity
from feed_pipeline.config.path import PATH_FEED_ORI, PATH_FEED_PROCESSED

def process_single_file(file_name):
    print(file_name)
    meta = FeedFileMeta(file_name)
    df = read_feed_excel(PATH_FEED_ORI / f'{file_name}.xlsx')
    df = calc_relative_age(df, start_date=meta.start_date)
    
    df_col, df_build = split_granularity(df)
    # print(1)
    save_incremental_parquet(
        df_col, PATH_FEED_PROCESSED / f'{meta.stage}{meta.unit}_column.parquet'
    )
    save_incremental_parquet(
        df_build, PATH_FEED_PROCESSED / f'{meta.stage}{meta.unit}_build.parquet'
    )