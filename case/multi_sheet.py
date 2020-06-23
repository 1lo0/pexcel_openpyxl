#-*- coding:utf-8 -*-
# author: donttouchkeyboard@gmail.com
# software: PyCharm

import openpyxl

from style.default import ExcelFontStyle
from core.base import TableCommon

if __name__ == '__main__':
    """
        一个表格，多个sheet   
    """
    # 创建一个Excel表格文件
    wb = openpyxl.Workbook()
    # 获取当前的默认sheet
    ws = wb.active
    # 默认sheet数据、样式、过滤等填充
    TableCommon.excel_write_common(wb=wb, ws=ws, sheet_name='Appche_ABC',
                             subject_list=['A', 'B', 'C'],
                             info_list=[
                                 ['Accumulo', 'Bahir', 'Cassandra'],
                                 ['ActiveMQ', 'Beam', 'Celix'],
                                 ['Airflow', 'Brooklyn', 'Clerezza'],
                                 ['Avro', 'BVal', 'CXF']
                             ],
                             # 不保存
                             save=False,
                             style=ExcelFontStyle.get_default_style(),
                             header_style=ExcelFontStyle.get_header_style(),
                             max_row=1,
                             max_col=3,
                             col_wide_map={1: 40, 2: 40, 3: 40})

    ws1 = wb.create_sheet()
    TableCommon.excel_write_common(wb=wb, ws=ws1, sheet_name='Appche_DEF',
                                   target_name='multi_sheet.xlsx',
                                   subject_list=['D', 'E', 'F'],
                                   info_list=[
                                       ['DataFu', 'Eagle', 'Felix'],
                                       ['DB', 'Empire-db', 'Fineract'],
                                       ['DeltaSpike', '：）', 'Flex'],
                                       ['Directory', '：）', 'Flink']
                                   ],
                                   # 保存表格
                                   save=True,
                                   style=ExcelFontStyle.get_default_style(),
                                   header_style=ExcelFontStyle.get_header_style(),
                                   max_row=1,
                                   max_col=3,
                                   col_wide_map={1: 40, 2: 40, 3: 40})