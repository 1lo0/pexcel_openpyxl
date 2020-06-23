#-*- coding:utf-8 -*-
# author: donttouchkeyboard@gmail.com
# software: PyCharm



import openpyxl


from openpyxl import Workbook



from core.base import TableCommon
from style.default import ExcelFontStyle
tb = TableCommon()

wb = Workbook()

ws = wb.active


tb.excel_common(wb, ws, sheet_name='Appche_ABC',  target_name='/tmp/pexcel_appche_project.xlsx',
                               subject_list=['A', 'B', 'C'], subject_list_row=1,
                                    subject_list_cloumn=1, info_list=[
                                    ['Accumulo', 'Bahir', 'Cassandra'],
                                    ['ActiveMQ', 'Beam', 'Celix'],
                                    ['Airflow', 'Brooklyn', 'Clerezza'],
                                    ['Avro', 'BVal', 'CXF']], info_list_row=2,
                                    info_list_cloumn=1,save=True, style=ExcelFontStyle.get_default_style(),
                               header_style=ExcelFontStyle.get_header_style(),
                               max_row=1,
                               max_col=3,
                                cell_value_map={'A1':60, 'A2': 80},
                               col_map={1:40,2:40,3: 40})



