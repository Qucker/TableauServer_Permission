import pymssql
import requests
from urllib import parse
import json
import csv
import random, string
import tableauserverclient as TSC
import datetime


# 连接数据库
def sqlcon():
    connect = pymssql.connect('192.168.90.34', 'sa', 'P@ssw0rd', 'Dev_Beta')  # 建立连接
    if connect:
        print("连接成功!")
        return connect


# 读取泛微apiJson数据
def fwapi():
    # FormData = {"username": 'WBS', "password": '1002AA10000000002BNO', "deptcode": 'WBXT001'}
    # url = "http://oa.sucgm.cn:8090/api/resource/getHrmRoleMembers"
    # 建立生产环境连接
    FormData = {"username": 'WBS', "password": '1002AA10000000002BNO', "deptcode": 'WBXT001'}
    url = 'https://43.254.222.12/api/resource/getHrmRoleMembers'
    data = parse.urlencode(FormData)
    HEADERS = {'Content-Type': 'application/x-www-form-urlencoded;charset=utf-8'}
    content = requests.post(url=url, headers=HEADERS, data=data, verify=False).text
    content = json.loads(content)
    return content


# 获取tableau server上的所有用户id，name
def tabuser(server, tableau_auth):
    tbuser = []
    with server.auth.sign_in(tableau_auth):
        all_users, pagination_item = server.users.get()
        # print(all_users)
        for user in all_users:
            tbuser.append([user.id, user.name])
        print(tbuser)
        return tbuser


# 获取tableau server上的所有群组id，name
def tabgroup(server, tableau_auth):
    tbgroup = []
    with server.auth.sign_in(tableau_auth):
        all_groups, pagination_item = server.groups.get()
        # print(all_groups)
        for group in all_groups:
            tbgroup.append([group.id, group.name])
        # print(tbgroup)
        return tbgroup


def getTwoDimensionListIndex(L, value):
    """获得二维列表某个值的一维索引值
    思想：先选出包含value值的一维列表，然后判断此一维列表在二维列表中的索引
    """
    data = [data for data in L if data[1] == value]  # data=[(53, 1016.1)]
    if data != []:
        index = L.index(data[0])
        return index
    else:
        return "不存在该值"


# 创建随机密码函数
def pwdcreate():
    src = string.ascii_letters + string.digits
    list_passwd_all = random.sample(src, 5)  # 从字母和数字中随机取5位
    list_passwd_all.extend(random.sample(string.digits, 1))  # 让密码中一定包含数字
    list_passwd_all.extend(random.sample(string.ascii_lowercase, 1))  # 让密码中一定包含小写字母
    list_passwd_all.extend(random.sample(string.ascii_uppercase, 1))  # 让密码中一定包含大写字母
    random.shuffle(list_passwd_all)  # 打乱列表顺序
    str_passwd = ''.join(list_passwd_all)  # 将列表转化为字符串
    return str_passwd


def main():
    '''主函数
    '''
    # 连接tableau server
    # tableau_auth = TSC.TableauAuth('20040201', 'GT2020!qaz', site_id='')
    tableau_auth = TSC.TableauAuth('GT_Account', 'GT2020123456', site_id='')
    server = TSC.Server('https://fas-gt.com:8000')
    # server = TSC.Server('https://tab.sucgm.com:10000')

    content = fwapi()
    infos = content['data']
    print(infos)
    connect = sqlcon()
    workgroup = []
    cursor = connect.cursor()  # 创建一个游标对象,python里的sql语句都要通过cursor来执行

    for info in infos:
        # 遍历获取员工信息号和群组名称
        workcode = info['workcode']
        # workcode = 'A0716'
        groupidnew = info['rolesmark']
        groupidnew = groupidnew[:-5]  # 将群组名称后的 -致同群组 字段删除
        workgroup.append([workcode, groupidnew])
        print(workcode)
        # 查询FW_Users库内人员状态信息
        sql = "select * from OFW_USER_TAB_ALLOWED where IsDelete = 0 and MemberCode = %s"
        cursor.execute(sql, workcode)  # 执行sql语句
        row = cursor.fetchall()  # 读取查询结果,
        rowCount = len(row)
        print(row)
        # for n in row:
        #     print(n[1])
        print(rowCount)
        tbgroup = tabgroup(server, tableau_auth)
        tbuser = tabuser(server, tableau_auth)
        # 查询用户是否已存在 若不存在 创建新用户并加入相应群组
        if rowCount == 0:
            # 查询该用户是否之前曾有过群组的记录
            sql = "select * from OFW_USER_TAB_ALLOWED where IsDelete = 0 and MemberCode = %s and GroupID= %s"
            value = (workcode, groupidnew)
            cursor.execute(sql, value)  # 执行sql语句
            row = cursor.fetchall()  # 读取查询结果,
            rowCount = len(row)
            if rowCount != 0:
                sql = "UPDATE OFW_USER_TAB_ALLOWED SET IsDelete =0 WHERE MemberCode = %s and GroupID= %s"
                cursor.execute(sql, value)
                connect.commit()

            else:
                sql2 = "insert into OFW_USER_TAB_ALLOWED (MemberCode,GroupID) VALUES (%s,%s)"
                value = (workcode, groupidnew)
                cursor.execute(sql2, value)
                connect.commit()
        # 用户已存在的情况
        else:
            # 查询该用户是否之前曾有过群组的记录
            sql = "select * from OFW_USER_TAB_ALLOWED where IsDelete = 1 and MemberCode = %s and GroupID= %s"
            value = (workcode, groupidnew)
            cursor.execute(sql, value)  # 执行sql语句
            rowold = cursor.fetchall()  # 读取查询结果,
            rowCount = len(rowold)
            if rowCount != 0:
                sql = "UPDATE OFW_USER_TAB_ALLOWED SET IsDelete =0 WHERE MemberCode = %s and GroupID= %s"
                cursor.execute(sql, value)
                connect.commit()
            else:
                oldgroup = []
                # 该员工原有的群组
                for n in row:
                    oldgroup.append(n[1])
                    # api传递的数据是一一对应的情况
                print(oldgroup)
                if groupidnew not in oldgroup:
                    sql2 = "insert into OFW_USER_TAB_ALLOWED (MemberCode,GroupID) VALUES (%s,%s)"
                    value = (workcode, groupidnew)
                    cursor.execute(sql2, value)
                    connect.commit()
    f.close()

    # cursor = connect.cursor()  # 创建一个游标对象,python里的sql语句都要通过cursor来执行
    # # 查询群组表内的群组信息
    # sql = "select GroupName from OFW_GROUP_TAB_ALLOWED"
    # cursor.execute(sql)  # 执行sql语句
    # rowgroup = cursor.fetchall()  # 读取查询结果,
    # # rowCount = len(row)
    # for groupid in rowgroup:
    #     # groupid = groupid[0]
    #     groupid = groupid[0]
    #     groupid = groupid[:-5] # 去除 -致同群组 字段
    #     # 查询表内对应该群组的workcode
    #     sqls = "select MemberCode from OFW_USER_TAB_ALLOWED_TEST where GroupID= %s"
    #     cursor.execute(sqls, groupid)  # 执行sql语句
    #     rowidingroups = cursor.fetchall()  # 读取查询结果,
    #     # print(rowidingroups)
    #     for rowidingroup in rowidingroups:
    #         rawdata = rowidingroup[0]
    #         # 此处注意判断是否有另外开的账号 需要加在指定群组中 但不在泛微数据中
    #         if [rawdata, groupid] not in workgroup:
    #             # 通过tsc 获取id 在group中删除
    #             # indexgroup=getTwoDimensionListIndex(tbgroup,groupid)
    #             indexuser = getTwoDimensionListIndex(tbuser, rawdata)
    #             with server.auth.sign_in(tableau_auth):
    #                 # mygroup = tbgroup[indexgroup][0]
    #                 userdel = tbuser[indexuser][0]
    #                 all_groups, pagination_item = server.groups.get()
    #                 for group in all_groups:
    #                     if group.name == groupid:
    #                         server.groups.remove_user(group, userdel)
    #                         # 若该群组不存在 进行逻辑删除
    #                         sql = "UPDATE OFW_USER_TAB_ALLOWED_TEST SET IsDelete =1 WHERE MemberCode = %s and GroupID= %s"
    #                         value = (rawdata, groupid)
    #                         cursor.execute(sql, value)
    #                         # 提交到数据库执行
    #                         connect.commit()
    #                 print('完成对' + rawdata + "在群组" + groupid + "的删除操作")
    #
    #                 # all_groups, pagination_item = server.groups.get()
    #                 # for group in all_groups:
    #                 #     if group.name == groupid:
    #                 #         mygroup = server.groups.get_by_id(group.id)
    #                 #         pagination_item = server.groups.populate_users(mygroup)
    #                 #         # print the names of the users
    #                 #         for user in mygroup.users:
    #                 #             if user.name == rawdata:
    #                 #                 server.groups.remove_user(mygroup, rawdata)


if __name__ == '__main__':
    main()
    print('End')
