import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
import pandas as pd

wb = pd.ExcelFile('Client Detail_2021_2H_concat_slim_pandas.xlsx', engine='openpyxl')
worksheets = wb.sheet_names # gets sheet names

# trying to print data for client
# client = 'Kammok'
# df = pd.read_excel(wb, 'FedEx', header=0)
# df = df.where(df['3PL Client']==client)
# df = df.dropna()
# print(df.head(10))

# gets clients from weekly pallet counts
wpc_df = pd.read_excel(wb, 'Weekly Pallet Counts')
clients = wpc_df['concat'].unique()
print(clients)

# for sheet in worksheets:
#     client = 'Airhouse'
#     df = pd.read_excel(spreadsheet, 'FedEx', header=0)
#     df = df.where(df['3PL Client']==client)
#     df = df.dropna()
#     print(df.head(10))

# ## pandas to create the files
# excel_file = 'ClientTest.xlsx'

# def open_excel(filename):
#     df = pd.read_excel(filename)
#     return df


# def split_data(raw_data):
#     df_processed = pd.DataFrame(raw_data)
#     headers = list(df_processed.columns.values)
#     customer_list = df_processed['3PL Client'].unique()

#     for index, row in df_processed.iterrows():
#         for i in customer_list:
#             df_edit = pd.DataFrame(row)
#             df_transposed = df_edit.T
#             df_transposed.to_excel(f'Output_Customer_{i}.xlsx', index=False, columns=headers)