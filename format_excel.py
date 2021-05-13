import shutil

import openpyxl
import os
import cpca

import datetime

if __name__ == '__main__':
    wb = openpyxl.load_workbook('target.xlsx')
    ws=wb.get_sheet_by_name('Sheet1')
    ws.delete_cols(17,3)
    del wb['过滤掉']
    wb.save("target.xlsx")

