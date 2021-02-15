import openpyxl
from openpyxl import Workbook, load_workbook
import pandas as pd
from pandas import ExcelWriter
from datetime import datetime

now = datetime.now()
dt_string = now.strftime("%d-%m-%Y_%H-%M-%S")

wb = pd.ExcelFile('Client Detail_2021_2H_concat.xlsx', engine='openpyxl')
worksheets = wb.sheet_names # gets sheet names - works but also creates sheet named ' (200)  Storage' that needs to be deleted
new_sheets = []

del worksheets[0]

# could probably use a list comprehension here
for sheet in worksheets:
    sheet_frame = pd.read_excel(wb, sheet, header=0)
    if 'concat' in sheet_frame.columns:
        new_sheets.append(sheet)

# gets clients from weekly pallet counts - works
wpc_df = pd.read_excel(wb, 'Weekly Pallet Counts')
clients = wpc_df['concat'].unique()


for client in clients:
    writer = ExcelWriter(f'{client}_Details_{dt_string}.xlsx')
    for sheet in new_sheets:
        sheet_frame = pd.read_excel(wb, sheet, header=0)
        if 'concat' in sheet_frame.columns:
            client_sheet = sheet_frame[sheet_frame['concat'] == client]
            if client_sheet.shape[0] == 0:
                continue
            # with pd.ExcelWriter(f'{client}_Details_{dt_string}.xlsx') as writer:  # pylint: disable=abstract-class-instantiated
            client_sheet.to_excel(writer, index=False, sheet_name=sheet)
        else:
            pass
    writer.save()
