import shutil

import openpyxl
import os
import cpca
from openpyxl.styles import numbers, is_date_format
from openpyxl import Workbook
import datetime

if __name__ == '__main__':
    sheet_name = '基本信息'
    # 创建导出表格
    if (not os.path.exists('target.xlsx')):
        shutil.copy('template.xlsx', 'target.xlsx')
    wb = openpyxl.load_workbook('target.xlsx')
    ws = wb.get_sheet_by_name('Sheet1')
    # 读取
    for root, dir, file in os.walk('sourceExcel'):
        for f in file:
            source_wb = openpyxl.load_workbook('sourceExcel/'+f)
            # 选择表单
            source_sh = source_wb.get_sheet_by_name(sheet_name)
            row_has = ws.max_row
            for i in range(1, source_sh.max_row - 2):
                ws.cell(i + row_has, 1, source_sh.cell(i + 1, 8).value)
                dict = { '招标公告': '公告', '招标预告': '预告','招标结果':'中标' }
                key=source_sh.cell(i + 1, 6).value
                ws.cell(i + row_has, 2, dict[key] if key in dict else key)
                ws.cell(i + row_has, 3, source_sh.cell(i + 1, 7).value)
                str = [source_sh.cell(i + 1, 5).value]
                df = cpca.transform(str)
                ws.cell(i + row_has, 7, df.iloc[0, 0])
                ws.cell(i + row_has, 8, df.iloc[0, 1])
                ws.cell(i + row_has, 9, source_sh.cell(i + 1, 9).value)
                ws.cell(i + row_has, 11, source_sh.cell(i + 1, 1).value)
                ws.cell(i + row_has, 12, '')
                ws.cell(i + row_has, 13, '')
                ws.cell(i + row_has, 14, '')
                dttm = datetime.datetime.strptime(source_sh.cell(i + 1, 2).value, "%Y-%m-%d")
                ws.cell(i + row_has, 15, dttm)
                ws.cell(i + row_has, 15).number_format = 'yyyy/m/d;@'
                ws.cell(i + row_has, 16, '')
            #
            wb.save("target.xlsx")
