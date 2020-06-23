# pexcel

## Introduction

```
基于openpyxl封装的上层应用，开箱即用。
只需要三行代码即可导出数据。
暂时只支持数据导出（对导入感兴趣的同学，请参考openpyxl官网）

暂时支持以下：
	自定义表格名称
        自定义sheet名称
        表头list动态填充
        数据list动态填充
        设置字体风格、表头背景色
        自定义行高、列宽
        支持排序、筛选
        友好的单元格合并功能
        上层单个Excel支持多个sheet
        上层可根据merge cell 定制两级、三级、多级复杂表头

```

## Install openpyxl

```
pip install openpyxl
```



## Qucik Start

> ![quickstart](https://gitee.com/WorldLine/pexecl/raw/master/images/pexcel_quickstart.png)

```python


import openpyxl

from style.default import ExcelFontStyle
from core.base import TableCommon


if __name__ == '__main__':
    """
        大小：
            导出一个5x4的表格
        表头：
            单行表头
            表格的表头为 A, B, C；表头文字格式加粗；
            位置居中；每列提供筛选、排序功能（默认开启，可设置sort_filter=false参数关闭）
        数据
            每一行为info_list的一个元素


        如果只是一个简单表格(单行表头)，你只需要了解一下参数
        TableCommon.excel_common
            :param workbook: excel
            :param worksheet: sheet
            :param sheet_name: sheet名称
            :param target_name: 保存目录的地址，默认在当前文件下
            :param subject_list: 表头信息  格式见default参数
            :param info_list: 数据集合   格式见default参数
            :param max_row: 表头所占行个数
            :param max_col: 表头所占列个数
            :return: Excel
    """
    # 创建一个Excel表格文件
    wb = openpyxl.Workbook()
    # 获取当前的默认sheet
    ws = wb.active
    # 默认sheet数据、样式、过滤等填充
    TableCommon.excel_common(wb=wb, ws=ws, sheet_name='Appche_ABC',
                             target_name='pexcel_appche_project.xlsx',
                             subject_list=['A', 'B', 'C'],
                             info_list=[
                                    ['Accumulo', 'Bahir', 'Cassandra'],
                                    ['ActiveMQ', 'Beam', 'Celix'],
                                    ['Airflow', 'Brooklyn', 'Clerezza'],
                                    ['Avro', 'BVal', 'CXF']
                                ],
                             save=True,
                             style=ExcelFontStyle.get_default_style(),
                             header_style=ExcelFontStyle.get_header_style(),
                             max_row=1,
                             max_col=3,
                             col_wide_map={1: 40, 2: 40, 3: 40})
```

## other

复杂表头请参考 `case package`


