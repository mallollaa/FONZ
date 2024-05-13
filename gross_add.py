
import pandas as pd
from utilities import calculate_slab, calculate_commission, add_total_row

def process_gross_add(df):
    df = df.dropna(axis=1, how='all')

    df = df[df['Reason'].isin(['Port in', 'Transfer Ownership', 'N/A'])]

    df['Package Price'] = df['Amount'] + df['DiscountAmount']
    df['Slab'] = df['Package Price'].apply(calculate_slab)
    df['Commission (KD)'] = calculate_commission(df['Package Price'], df['Slab'])
    df = add_total_row(df, 'Commission (KD)')

    return df

if __name__ == "__main__":
    source_file = 'path_to_your_excel_file.xlsx'
    df = pd.read_excel(source_file, sheet_name='Gross Add')
    processed_df = process_gross_add(df)
    processed_df.to_excel('output_gross_add.xlsx', index=False)
