import os
import openpyxl
import pandas as pd

'''
in_wb = openpyxl.load_workbook('input/angaza_payment.xlsx')
in_sheet = in_wb['ap']

max_row = in_sheet.max_row
max_col = in_sheet.max_column

row_no = 2

print(in_sheet)
'''

df = pd.read_excel('input/angaza_payment.xlsx')
print(df['E'].unique())

