# -*- coding:utf-8 -*-
# author: donttouchkeyboard@gmail.com
# software: PyCharm

import openpyxl

from style.default import ExcelFontStyle
from core.base import TableCommon

if __name__ == '__main__':

    """
        多级复杂表头绘制
            merge_map：单元格合并
            merge_cell_value_map：合并单元格赋值
    """

    wb = openpyxl.Workbook()
    ws = wb.active
    wb = TableCommon.excel_write_common(
        wb=wb,
        ws=ws,
        sheet_name='Beijing_10subway',
        target_name='pexcel_beijing_10subway.xlsx',
        subject_list=[
            '巴沟站', '苏州街站', '海淀黄庄站', '知春里站', '知春路站', '西土城站', '牡丹园站', '健德门站', '北土城站',
            '安贞门站', '惠新西街南口站', '芍药居站', '太阳宫站', '三元桥站', '亮马桥站', '农业展览馆站', '团结湖站', '呼家楼站',
            '金台夕照站', '国贸站'
        ],
        subject_list_row=3,
        subject_list_cloumn=4,
        info_list=[
            ['20190911', 90, 22, 90, 22, 90, 90, 22, 90, 22, 90, 90, 22, 90, 22, 90, 90, 22, 90, 22,
             90, 22, 90],
            ['20190912', 90, 22, 90, 22, 90, 90, 22, 90, 22, 90, 90, 22, 90, 22, 90, 90, 22, 90, 22,
             90, 22, 90],
            ['20190913', 90, 22, 90, 22, 90, 90, 22, 90, 22, 90, 90, 22, 90, 22, 90, 90, 22, 90, 22,
             90, 22, 90],
            ['20190914', 90, 22, 90, 22, 90, 90, 22, 90, 22, 90, 90, 22, 90, 22, 90, 90, 22, 90, 22,
             90, 22, 90],
            ['20190915', 90, 22, 90, 22, 90, 90, 22, 90, 22, 90, 90, 22, 90, 22, 90, 90, 22, 90, 22,
             90, 22, 90],
            ['20190916', 90, 22, 90, 22, 90, 90, 22, 90, 22, 90, 90, 22, 90, 22, 90, 90, 22, 90, 22,
             90, 22, 90],
        ],
        merge_map={'A1': 'A3', 'B1': 'E1', 'B2': 'B3', 'C2': 'C3',
                   'D2': 'E2', 'F1': 'L1', 'F2': 'I2', 'J2': 'L2',
                   'M1': 'W1', 'M2':'O2', 'P2':'R2', 'S2':'W2'},

        merge_cell_value_map={'A1': 'Date', 'B1': 'AA',
                              'B2': 'DD', 'C2': 'EE', 'D2': 'aa',
                              'F1': 'BB', 'F2': 'bb', 'J2': 'cc',
                              'M1': 'CC', 'M2': 'dd', 'P2': 'ee',
                              'S2': 'ff'
                              },
        info_list_row=4,
        info_list_cloumn=1,
        save=True,
        style=ExcelFontStyle.get_default_style(),
        max_row=3,
        max_col=23,
        col_wide_map={1: 20, 22: 40, 21: 40})
