import pandas as pd
import openpyxl
from main import kam_df
try:
    df = pd.merge(df, kam_df[['ID', 'Name']], left_on='PayeeID', right_on='ID', how='left')
    df.drop(columns=['ID'], inplace=True)
    df['New PayeeID'] = df['Name'].fillna('FONZ')
    df.drop(columns=['Name'], inplace=True)
except KeyError as e:
    print(f"Error merging data: {e}")
    print("Check column names in the source and KAM data files.")
    exit(1)

print("KAM DataFrame columns:", kam_df.columns)
