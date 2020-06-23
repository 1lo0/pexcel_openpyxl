#-*- coding:utf-8 -*-
# author: donttouchkeyboard@gmail.com
# software: PyCharm
import openpyxl

from core.base import TableCommon
from style.default import ExcelFontStyle

class StartPexcel(TableCommon):
    """
    base openpyxl excel util

    pexecel hello world

    include 1 excel file with 2 sheet
    """

    @classmethod
    def hello_pexcel_appche_abc(self,
                                wb,
                                ws,
                                sheet_name='Appche_ABC',
                                target_name='test_merge_opt.xlsx',
                                subject_list=['A', 'B', 'C'],
                                info_list=[
                                    ['Accumulo', 'Bahir', 'Cassandra'],
                                    ['ActiveMQ', 'Beam', 'Celix'],
                                    ['Airflow', 'Brooklyn', 'Clerezza'],
                                    ['Avro', 'BVal', 'CXF']
                                ],
                                save=False,
                                style=ExcelFontStyle.get_default_style(),
                                ):
        """

        :param wb: excel
        :param ws: sheet
        :param sheet_name: sheet名称
        :param target_name: 保存目录的地址，默认在当前文件下
        :param subject_list: 表头信息  格式见default参数
        :param info_list: 数据集合   格式见default参数
        :param save: 是否保存 默认false
        :param style: 单元格字体、背景
        :other 如果存在复杂表头需定制。     复杂表头可参考 case 模块
        :return: Excel
        """
        wb = self.excel_common(wb, ws, sheet_name, target_name,
                               subject_list, 1, 1, info_list, 2, 1, save, style,
                               header_style=ExcelFontStyle.get_header_style(),
                               max_row=1,
                               max_col=3,
                               merge_map={'B4':'C5', 'A4':'A5'},
                               col_map={1:40, 2:40, 3: 40},
                               cell_value_map={'2-1':6666666666}
                               )
        return wb


if __name__ == '__main__':
    # 创建一个Excel表格文件
    wb = openpyxl.Workbook()
    # 获取当前的默认sheet
    ws = wb.active
    # 默认sheet数据、样式、过滤等填充
    wb = StartPexcel.hello_pexcel_appche_abc(wb=wb, ws=ws, save=True)
