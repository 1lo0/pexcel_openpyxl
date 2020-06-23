# -*- coding:utf-8 -*-
# author: donttouchkeyboard@gmail.com
# software: PyCharm
import openpyxl

from style.default import ExcelFontStyle
from core.base import TableCommon

if __name__ == '__main__':

    """
        二级表头示例
    """

    wb = openpyxl.Workbook()
    ws = wb.active
    wb = TableCommon.excel_write_common(wb, ws, sheet_name='Secondary_Header',
                                  target_name='pexcel_secondary_header.xlsx',
                                  subject_list=['prd', 'onl', 'prd', 'onl', 'prd', 'onl', 'prd', 'onl'],
                                  subject_list_row=2, subject_list_cloumn=2,
                                  info_list=[
                                      ['2019.3', 22, 90, 22, 90, 22, 90, 22, 90],
                                      ['2019.4', 22, 90, 22, 90, 22, 90, 22, 90],
                                      ['2019.5', 22, 90, 22, 90, 22, 90, 22, 90],
                                      ['2019.6', 22, 90, 22, 90, 22, 90, 22, 90],
                                  ],
                                  merge_map={'A1': 'A2', 'B1': 'C1', 'D1': 'E1', 'F1': 'G1', 'H1': 'I1'},
                                  merge_cell_value_map={'A1': 'Month', 'B1': 'AA',
                                                        'D1': 'BB', 'F1': 'CC', 'H1': 'DD'},
                                  info_list_row=3,
                                  info_list_cloumn=1,
                                  save=True,
                                  style=ExcelFontStyle.get_default_style(),
                                  header_style=ExcelFontStyle.get_header_style(),
                                  max_row=2,
                                  max_col=9,
                                  col_wide_map={1: 13})
