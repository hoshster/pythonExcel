#!/usr/local/bin/python3

import os
import openpyxl

print(os.getcwd())
os.chdir(os.path.join(os.path.expanduser('~'), 'Desktop'))

if not os.path.exists('FlowedFileReport'):
    os.mkdir('FlowedFileReport')
else:
    pass

os.chdir('FlowedFileReport')

# open workbook
wb = openpyxl.load_workbook("report.xlsx")
wb.active()
sh = wb['Sheet1']
type(sh)


