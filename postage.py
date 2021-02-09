import pandas as pd
import numpy as np
from datetime import datetime

now = datetime.now()
dt_string = now.strftime("%d-%m-%Y_%H-%M-%S")

postage = pd.read_excel('Postage_stripped.xlsx', sheet_name='Endicia')
shipments = pd.read_excel('Shipments.xlsx', sheet_name='SH_Shipments')

# Postage BA file is col 53. Is inserted into 54
# Shipments searching col 11, returning what's in column 17

new_postage = postage.merge(shipments[['Tracking Number', '3PL Customer']], on='Tracking Number')
new_postage.to_excel(f'testPostage_{dt_string}.xlsx')