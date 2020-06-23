#-*- coding:utf-8 -*-
# author: donttouchkeyboard@gmail.com
# software: PyCharm


import datetime

def get_concat_shell():
    """

    :return:
    """

    gen_shell = """
    sudo -u hdfs hadoop distcp   hdfs://192.168.60.121:8020/data/user/hive/warehouse/cnk.db/dwd_cnk_nginx_base/dt={} hdfs://192.168.60.116:8020/data/user/hive/warehouse/dwd_kylin.db/dwd_kylin_nginx_base/
    """

    import datetime
    d2 = datetime.datetime(2020, 2, 2)
    d1 = datetime.datetime(2020, 4, 1)
    xDay = (d1 - d2).days
    aDay = datetime.timedelta(days=1)
    i = 0
    while i < xDay:
        i += 1
        d2 += aDay
        print(gen_shell.format(d2.strftime('%Y-%m-%d')))


if __name__ == '__main__':
    get_concat_shell()