from feed_analysis.feed_pipeline.utils.process_single_file import process_single_file
from feed_analysis.config.path import PATH_FEED_ORI

PATH = PATH_FEED_ORI

def main():
    # process_single_file('2023-01-01')
    for file in PATH.glob('*.xlsx'):
        process_single_file(file.stem)

if __name__ == '__main__':
    main()  

    # Terminal: python -m feed_analysis.main