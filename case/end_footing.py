# -*- coding:utf-8 -*-
# author: donttouchkeyboard@gmail.com
# software: PyCharm


import openpyxl

from style.default import ExcelFontStyle
from core.base import TableCommon

if __name__ == '__main__':
    """
        在表格后最后一列追加合计、统计信息(append_list参数)
            ps: 也可以放到info_list最后一个元素实现
    """

    wb = openpyxl.Workbook()
    ws = wb.active
    wb = TableCommon.excel_write_common(
        wb=wb,
        ws=ws,
        sheet_name='Social_Contact_App',
        target_name='pexcel_social_contact_app.xlsx',
        subject_list=[
            '微信', 'QQ', '微博', '百度贴吧', '豆瓣', '人人',
            '微信', 'QQ', '微博', '百度贴吧', '豆瓣', '人人'
        ],
        subject_list_row=2,
        subject_list_cloumn=2,
        info_list=[
            ['2019.5', 23, 22, 90, 22, 90, 22, 90, 22, 90, 80, 90, 80],
            ['2019.6', 23, 22, 90, 22, 90, 22, 90, 22, 90, 80, 90, 80],
            ['2019.7', 23, 22, 90, 22, 90, 22, 90, 22, 90, 80, 90, 80],
            ['2019.8', 23, 22, 90, 22, 90, 22, 90, 22, 90, 80, 90, 80],
        ],
        merge_map={'A1':'A2', 'B1':'G1', 'H1':'m1'},
        merge_cell_value_map={'A1':'Date', 'B1':'A', 'H1':'B'},
        info_list_row=3,
        info_list_cloumn=1,
        save=True,
        style=ExcelFontStyle.get_default_style(),
        append_list=[['合计', 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100]],
        max_row=2,
        max_col=13,
        col_wide_map={1: 20}
    )
