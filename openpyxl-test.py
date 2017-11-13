# -*- coding: utf-8 -*-
"""
Created on Wed Nov  8 16:53:56 2017

@author: Frank
"""

from openpyxl import load_workbook
wb = load_workbook(filename='test.xlsx')
ws = wb.get_sheet_by_name(name='Sheet1')
columns = ws.columns()
#print(columns[0])