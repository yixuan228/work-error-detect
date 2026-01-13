def calc_relative_age(df, start_date):
    df['age'] = [(d - start_date).days for d in df['Date']]
    return df

import pandas as pd
def recalibrate_age(df, reference_date, reference_age):
    df['age'] = [(d - pd.to_datetime(reference_date).date()).days + reference_age for d in df['Date']]  
    return df
