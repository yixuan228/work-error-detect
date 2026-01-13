def calc_relative_age(df, start_date):
    df['age'] = [(d - start_date).days for d in df['Date']]
    return df


def recalibrate_age(df, reference_date, reference_age):
    df['age'] = [(d - reference_date).days + reference_age for d in df['Date']]
    return df
