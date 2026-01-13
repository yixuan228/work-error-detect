import pandas as pd

def save_incremental_parquet(df, file_path):
    if file_path.exists():
        prev = pd.read_parquet(file_path, engine='pyarrow')
        df = (
            pd.concat([prev, df], ignore_index=True)
            .drop_duplicates()
            .sort_values(by='Date')
        )
    df.to_parquet(file_path, index=False)