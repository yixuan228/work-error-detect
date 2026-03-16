import pandas as pd

def save_incremental_parquet(df, file_path):
    """
    增量保存数据
    """
    if file_path.exists():
        prev = pd.read_parquet(file_path, engine='pyarrow')
        df = (
            pd.concat([prev, df], ignore_index=True)
            .drop_duplicates()
            
        )
        sort_cols = ["Date", "col_num"]
        exist_cols = [c for c in sort_cols if c in df.columns]
        df['Date'] = pd.to_datetime(df['Date']).dt.date
        df = df.sort_values(by=exist_cols).reset_index(drop=True)
    df.to_parquet(file_path, index=False)
