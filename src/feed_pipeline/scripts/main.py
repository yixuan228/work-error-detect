from feed_pipeline.scripts.batch_process import process_single_file
from feed_pipeline.config.path import PATH_FEED_ORI

PATH = PATH_FEED_ORI

for file in PATH.glob('*.xlsx'):
    process_single_file(file.stem)