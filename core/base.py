# -*- coding:utf-8 -*-
# author:donttouchkeyboard@gmail.com
# software: PyCharm


from style.default import ExcelFontStyle
from common.util import ExcelUtil
from common.val import sort_filter_base


class TableCommon:
    """
    表格common功能点：
        自定义表格名称
        自定义sheet名称
        表头list动态填充
        数据list动态填充
        设置字体风格、表头背景色
        自定义行高、列宽
        支持排序、筛选
        上层单个Excel支持多个sheet
        上层可根据merge cell 定制复杂表头
    """

    @classmethod
    def excel_write_common(self,
                     wb,
                     ws,
                     sheet_name='',
                     target_name='',
                     subject_list=[],
                     subject_list_row=1,
                     subject_list_cloumn=1,
                     info_list=[],
                     info_list_row=2,
                     info_list_cloumn=1,
                     save=False,
                     style=ExcelFontStyle.get_default_style(),
                     append_list=[],
                     header_style=ExcelFontStyle.get_header_style(),
                     max_row=0,
                     max_col=0,
                     row_high_map={},
                     col_wide_map={},
                     merge_map={},
                     merge_cell_value_map={},
                     sort_filter=True,
                     ):
        """
        excel common
        :param wb: excel
        :param ws: sheet
        :param sheet_name: sheet名称
        :param target_name: 宿主机保存目录地址
        :param subject_list: 表头信息
        :param subject_list_row: subject_list 起始行    默认1
        :param subject_list_cloumn: subject_list 起始列 默认1
        :param info_list: 数据集合信息
        :param info_list_row: info_list 起始行   默认2
        :param info_list_cloumn: info_list 起始列    默认1
        :param save: 是否保存 默认false
        :param style: 单元格字体、背景
        :param append_list: 最后追加list, 一般适用于合计
        :param header_style: 表头字体、背景
        :param max_row: 表头所占行个数
        :param max_col: 表头所占列个数
        :param row_high_map: 自定义行高  eg:{1:10, 2:20} 第一行的行高为10  第二行的行高为20
        :param col_wide_map: 自定义列宽 eg:{1:30, 2:40} 第一列的列宽为30， 第二列的列宽为40
        :param merge_map: 自定义合并单元格范围   eg: {'B4':'C5', 'A4':'A5'}
        :param merge_cell_value_map: 自定义cell单元格的value  eg: {'A1': pexcel, 'A2': python}
        :param sort_filter: 数据支持排序（默认支持，不需要可设为 false）
        """

        # sheet名称
        ws.title = sheet_name

        # 信息列表
        if subject_list.__len__() > 0:
            for i in range(subject_list.__len__()):
                ws.cell(row=subject_list_row, column=i + subject_list_cloumn, value=subject_list[i])
        # 列长度
        if info_list.__len__() > 0:
            cloums = info_list[0].__len__()
            # 写入信息
            for obs in range(info_list.__len__()):
                for info in range(cloums):
                    ws.cell(row=obs + info_list_row, column=info + info_list_cloumn, value=info_list[obs][info])

        # append list
        if append_list.__len__() > 0:
            for record in append_list:
                ws.append(record)

        # 数据格式应用
        for row in ws.rows:
            for cell in row:
                cell.style = style

        # 设置表头对齐方式、背景色、字体大小
        for eachCommonRow in ws.iter_rows(max_row=max_row, max_col=max_col):
            for eachCellInRow in eachCommonRow:
                eachCellInRow.alignment = header_style.alignment
                eachCellInRow.fill = header_style.fill
                eachCellInRow.font = header_style.font

        # 设置header行的默认高度
        for i in range(1, max_row + 1):
            ws.row_dimensions[i].height = 40
        # 设置header列的默认宽度
        for i in range(1, max_col + 1):
            ws.column_dimensions[ExcelUtil.get_letter_list()[i]].width = 40.0
        # 设置行高，key：具体row value: 行高
        if row_high_map.__len__() > 0:
            for i in range(1, max_row + 1):
                row_height = row_high_map.get(i)
                ws.row_dimensions[i].height = row_height
        # 设置列宽，key：具体列 value: 列宽
        if col_wide_map.__len__() > 0:
            for i in range(1, max_col + 1):
                if col_wide_map.__contains__(i):
                    col_width = col_wide_map.get(i)
                    ws.column_dimensions[ExcelUtil.get_letter_list()[i - 1]].width = col_width
        # 设置单元格合并
        if merge_map.__len__() > 0:
            for mer in merge_map:
                ws.merge_cells(mer + ':' + merge_map.get(mer))
        # 设置合并后单元格的值
        if merge_cell_value_map.__len__() > 0:
            for cel in merge_cell_value_map:
                ws[cel] = merge_cell_value_map.get(cel)

        # 开启过滤，排序
        sort_filter_start = sort_filter_base + str(max_row)
        sort_filter_end = ExcelUtil.get_letter_list(True)[max_col - 1] + "" + str(max_row)
        if sort_filter:
            ws.auto_filter.ref = sort_filter_start + ":" + sort_filter_end
        # 保存文件
        if (save):
            wb.save(target_name)
        return wb
