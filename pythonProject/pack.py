from openpyxl import Workbook
import pandas as pd
import pprint



wb = Workbook()

# grab the active worksheet
ws = wb.active

order = pd.read_excel("test.xlsx", header=0)
panel=pd.read_excel("PanelsF.xlsx", header=0)


#order.loc[order['Type'] == 'PR-5T']

#take the unique coloumn or unige type in order
order_type=order['Type'].unique().tolist()
for i in order_type:
    print(i[:2])
    type=i[:2]
    print(order.loc[order['Type'] == i])


