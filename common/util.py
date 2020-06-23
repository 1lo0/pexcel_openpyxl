#-*- coding:utf-8 -*-
# author:donttouchkeyboard@gmail.com
# software: PyCharm





class ExcelUtil:
    """
    excel 工具类
    """

    @classmethod
    def get_letter_list(self, uppered = True):
        """
        获取26英文字母集合
        :param uppered: 默认大写
        :return: 26英文字母list
        """
        lower_case_list = [chr(i)  for i in range(97, 123)]
        if(uppered):
            upper_case_list = []
            for i in lower_case_list:
               upper_case_list.append(i.upper())
            return upper_case_list
        return lower_case_list

    def tuple_empty(self):
        """
        判断tuple是否为空
        :return:
        """





