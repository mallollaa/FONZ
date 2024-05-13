import pandas as pd
import os
import shutil
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo


source_file = 'Inputs/D2D_April.xlsx'
kam_file = 'Inputs/KAM_NAMES.xlsx'
WhereToBeSaved = '/Users/manalalajmi/PycharmProjects/FonzBOT/FonzOUTPUT'
new_file_name = 'Fonz_Commission_Report_Processed.xlsx'
saved_file = os.path.join(WhereToBeSaved, new_file_name)
os.makedirs(WhereToBeSaved, exist_ok=True)

shutil.copy(source_file, saved_file)

# KAMS
kam_df = pd.read_excel(kam_file, sheet_name='KAM')
kam_df.columns = kam_df.columns.str.strip()

wb = load_workbook(saved_file)
ws = wb.active

df = pd.read_excel(saved_file, sheet_name='sheet1')

df = df[df['PayeeID'] == 'FONZ']

def calculate_slab(package_price):
    if 0 <= package_price <= 6:
        return 1.75
    elif 6 < package_price <= 15:
        return 2
    elif package_price > 15:
        return 2.5
    return 0

def calculate_commission(package_price, slab):
    return package_price * slab

event_types = ["ACTIVATION", "DEACTIVATION", "GROSS ADD", "CHURN", "UPGRADE", "DOWNGRADE", "RENEW", "DEVICE ADD", "CLAWBACK"]
for event_type in event_types:
    event_df = df[df['EventType'].str.upper() == event_type].copy()
    if event_df.empty:
        continue
        # I will add the other event type later
    if event_type == "GROSS ADD":
        event_df = event_df[~event_df['Reason'].isin(['Port in', 'Transfer Ownership', 'N/A'])]
        event_df['Package Price'] = event_df['Amount'] + event_df['DiscountAmount']
        event_df['Slab'] = event_df['Package Price'].apply(calculate_slab)
        event_df['Commission (KD)'] = event_df['Package Price'] * event_df['Slab']
        columns_to_delete = ['OrderNumber', 'PriceType', 'PenaltyPeriod', 'Quantity', 'OfferType', 'MarketingCategory',
                             'TransactionDescription']
        event_df.drop(columns=[col for col in columns_to_delete if col in event_df.columns], inplace=True)


        total_commission = event_df['Commission (KD)'].sum()
        total_row = pd.DataFrame([{col: "" for col in event_df.columns}, {'Commission (KD)': total_commission}],
                                 index=['', 'Total'])
        event_df = pd.concat([event_df, total_row], ignore_index=True)

    if event_type not in wb.sheetnames:
        wb.create_sheet(event_type)
    event_sheet = wb[event_type]

    #This secion for Headers and styleee
    for r_idx, row in enumerate(dataframe_to_rows(event_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            event_sheet.cell(row=r_idx, column=c_idx, value=value)
    tab = Table(displayName=event_type.replace(" ", ""), ref="A1:" + get_column_letter(len(event_df.columns)) + str(len(event_df)+1))
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    event_sheet.add_table(tab)

wb.save(saved_file)
print("Yayyy Doneeee", saved_file)
