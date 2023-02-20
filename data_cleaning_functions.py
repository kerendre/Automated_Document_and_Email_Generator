import pandas as pd





def format_dates_2_d_m_Y(df, column_name):
    # Convert the specified column to a datetime object if it's not already in that format
    if not pd.api.types.is_datetime64_any_dtype(df[column_name]):
        df[column_name] = pd.to_datetime(df[column_name])

    # Convert the specified column to a string in the desired format "%d/%m/%Y"
    df[column_name] = df[column_name].dt.strftime("%d/%m/%Y")

    # Return the modified dataframe
    return df



def merge_data(df, df_contact,on_column):
    # Merge the two dataframes on the "municipal" column
    merged_df = pd.merge(df, df_contact, on=on_column)

    # Return the merged dataframe
    return merged_df


