import pandas as pd

def save_incremental_parquet(df, file_path):
    if file_path.exists():
        prev = pd.read_parquet(file_path, engine='pyarrow')
        df = (
            pd.concat([prev, df], ignore_index=True)
            .drop_duplicates()
            
        )
        df['Date'] = pd.to_datetime(df['Date']).dt.date
        df= df.sort_values(by='Date').reset_index(drop=True)
    df.to_parquet(file_path, index=False)