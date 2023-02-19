
def filter_df(df):
    # Get the list of column names that don't end in "heb" or "heb "
    df1_columns = [col for col in df.columns if not col.endswith('heb') and not col.endswith('heb ')]

    # Create a new DataFrame with the selected columns
    df1 = df[df1_columns]

    # Return the new DataFrame
    return df1
