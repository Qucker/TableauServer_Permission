import pymssql
import requests
from urllib import parse
import json
import csv
import random, string
import tableauserverclient as TSC
import datetime
import xlrd
import xlwt


def pwdcreate():
    src = string.ascii_letters + string.digits + string.punctuation
    list_passwd_all = random.sample(src, 6)  # 从字母,数字,特殊字符中随机取6位
    list_passwd_all.extend(random.sample(string.digits, 1))  # 让密码中一定包含数字
    list_passwd_all.extend(random.sample(string.ascii_lowercase, 1))  # 让密码中一定包含小写字母
    list_passwd_all.extend(random.sample(string.ascii_uppercase, 1))  # 让密码中一定包含大写字母
    list_passwd_all.extend(random.sample(string.punctuation, 1))  # 让密码中一定包含特殊字符
    random.shuffle(list_passwd_all)  # 打乱列表顺序
    str_passwd = ''.join(list_passwd_all)  # 将列表转化为字符串
    return str_passwd


def readExcel():
    workbook = xlrd.open_workbook(
        r'D:\BaiduNetdiskDownload\工作\城建市政\Tableau Server用户管理\TableauUserpwdChange_20201019.xlsx')
    sheet = workbook.sheets()[0]
    clou = sheet.col_values(0)  # 读取第一列
    return clou[1:]


def main():
    request_options = TSC.RequestOptions(pagesize=1000)
    tableau_auth = TSC.TableauAuth('20040201', 'GT2020!qaz', site_id='')
    # tableau_auth = TSC.TableauAuth('GT_Account', 'GT2020123456', site_id='')
    # server = TSC.Server('https://fas-gt.com:8000')
    server = TSC.Server('https://tab.sucgm.com:10000', use_server_version=True)
    f = open('user&pwd.csv', 'w', encoding='utf-8', newline='' "")
    # 2. 基于文件对象构建 csv写入对象
    csv_writer = csv.writer(f)

    # 3. 构建列表头
    csv_writer.writerow(["Workcode", "Password"])

    column = readExcel()

    with server.auth.sign_in(tableau_auth):
        for workcode in column:

            # 将新用户添加至user表
            all_users, pagination_item = server.users.get(request_options)
            for user in all_users:
                if user.name == workcode:
                    pwd = pwdcreate()
                    print(pwd)
                    user1 = server.users.update(user, password=pwd)
                    csv_writer.writerow([workcode, pwd])
            # else:
            #     print(index)

    f.close()


if __name__ == '__main__':
    main()
