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
    FormData = {"username": 'WBS', "password": '1002AA10000000002BNO', "deptcode": 'WBXT001'}
    data = parse.urlencode(FormData)
    url = "https://cp.sucgm.com/api/resource/getHrmRoleMembers"
    HEADERS = {'Content-Type': 'application/x-www-form-urlencoded;charset=utf-8'}
    content = requests.post(url=url, headers=HEADERS, data=data).text
    content = json.loads(content)
    return content


# 获取tableau server上的所有用户id，name
def tabuser(server, tableau_auth, request_options):
    tbuser = []
    with server.auth.sign_in(tableau_auth):
        all_users, pagination_item = server.users.get(request_options)
        # print(all_users)
        for user in all_users:
            tbuser.append([user.id, user.name])
        print(tbuser)
        return tbuser


# 获取tableau server上的所有群组id，name
def tabgroup(server, tableau_auth, request_options):
    tbgroup = []
    with server.auth.sign_in(tableau_auth):
        all_groups, pagination_item = server.groups.get(request_options)
        # print(all_groups)
        for group in all_groups:
            tbgroup.append([group.id, group.name])
        # print(tbgroup)
        return tbgroup


def getTwoDimensionListIndex(L, value):
    """获得二维列表某个值的一维索引值
    思想：先选出包含value值的一维列表，然后判断此一维列表在二维列表中的索引
    """
    data = [data for data in L if data[1] == value]
    if data != []:
        index = int(L.index(data[0]))
        return index
    else:
        return "不存在该值"


# 创建随机密码函数
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


def main():
    '''主函数
    '''
    # 连接tableau server
    request_options = TSC.RequestOptions(pagesize=200)

    # tableau_auth = TSC.TableauAuth('20040201', 'GT2020!qaz', site_id='')
    tableau_auth = TSC.TableauAuth('GT_Account', 'GT2020123456', site_id='')
    server = TSC.Server('https://fas-gt.com:8000', use_server_version=True)
    # server = TSC.Server('https://tab.sucgm.com:10000',use_server_version=True)

    nowTime = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')  # 现在
    content = fwapi()
    infos = content['data']
    # print("输出本次全部数据:",infos)
    connect = sqlcon()
    workgroup = []
    cursor = connect.cursor()  # 创建一个游标对象,python里的sql语句都要通过cursor来执行

    # 1. 创建文件对象
    f = open('user&pwd-' + nowTime + '.csv', 'w', encoding='utf-8', newline='' "")
    # 2. 基于文件对象构建 csv写入对象
    csv_writer = csv.writer(f)

    # 3. 构建列表头
    csv_writer.writerow(["Workcode", "Password"])
    tbgroup = tabgroup(server, tableau_auth, request_options)
    tbuser = tabuser(server, tableau_auth, request_options)

    # 本次离职人数计数器
    quitCount =0

    # 系统账户（不可删除）
    SystemUser=['SHCJTS_Admin','20040201','RPA_Account']
    SystemUsertest=['DA_Admin','GT_Account','Publishtester','Viewer']



    for info in infos:
        # 遍历获取员工信息号和群组名称
        workcode = info['workcode']
        # workcode = 'A0716'
        groupidnew = info['rolesmark']
        status = info['status']
        if status == '1':
            groupidnew = groupidnew[:-5]  # 将群组名称后的 -致同群组 字段删除
            workgroup.append([workcode, groupidnew])
            print(workcode)
            # 查询FW_Users库内对应workcoode 人员状态信息
            # sql = "select * from OFW_USER_TAB_ALLOWED where IsDelete = 0 and MemberCode = %s"
            sql = "select * from OFW_USER_TAB_ALLOWED_TEST where IsDelete = 0 and MemberCode = %s"
            cursor.execute(sql, workcode)  # 执行sql语句
            row = cursor.fetchall()  # 读取查询结果,
            rowCount = len(row)
            print(row)
            # for n in row:
            #     print(n[1])
            print(workcode, "该员工当前在Tableau Server对应群组记录数", rowCount)
            # time.sleep(5)

            # 查询用户是否已存在 若不存在 创建新用户并加入相应群组
            if rowCount == 0:
                # 查询该用户是否之前曾有过群组的记录
                # sql = "select * from OFW_USER_TAB_ALLOWED where IsDelete = 0 and MemberCode = %s and GroupID= %s"
                sql = "select * from OFW_USER_TAB_ALLOWED_TEST where IsDelete = 0 and MemberCode = %s and GroupID= %s"
                value = (workcode, groupidnew)
                cursor.execute(sql, value)  # 执行sql语句
                row = cursor.fetchall()  # 读取查询结果,
                rowCount = len(row)
                if rowCount != 0:
                    # sql = "UPDATE OFW_USER_TAB_ALLOWED SET IsDelete =0 WHERE MemberCode = %s and GroupID= %s"
                    sql = "UPDATE OFW_USER_TAB_ALLOWED_TEST SET IsDelete =0 WHERE MemberCode = %s and GroupID= %s"
                    cursor.execute(sql, value)
                    connect.commit()
                    with server.auth.sign_in(tableau_auth):
                        # 将用户加入群组
                        # index = getTwoDimensionListIndex(tbgroup, groupidnew)
                        # 根据workcode 取用户对应的tableau server id
                        user = getTwoDimensionListIndex(tbuser, workcode)
                        # print(user)
                        all_groups, pagination_item = server.groups.get(request_options)
                        for group in all_groups:
                            if group.name == groupidnew:
                                server.groups.add_user(group, tbuser[user][0])
                                break
                else:
                    with server.auth.sign_in(tableau_auth):
                        # 将新用户添加至user表
                        newU = TSC.UserItem(workcode, 'Viewer')
                        # add the new user to the site
                        newU = server.users.add(newU)
                        # 将用户的id，name加入tbuser
                        tbuser.append([newU.id, newU.name])
                        # 更新详细信息
                        newU.fullname = workcode
                        # index = getTwoDimensionListIndex(tbgroup, groupidnew)

                        # 加入群组
                        # if index !="不存在该值":
                        all_groups, pagination_item = server.groups.get(request_options)
                        for group in all_groups:
                            if group.name == groupidnew:
                                server.groups.add_user(group, newU.id)
                                break
                        # else:
                        #     print(index)
                        pwd = pwdcreate()
                        csv_writer.writerow([workcode, pwd])
                        user1 = server.users.update(newU, password=pwd)
                    # sql2 = "insert into OFW_USER_TAB_ALLOWED (MemberCode,GroupID) VALUES (%s,%s)"
                    sql2 = "insert into OFW_USER_TAB_ALLOWED_TEST (MemberCode,GroupID) VALUES (%s,%s)"
                    value = (workcode, groupidnew)
                    cursor.execute(sql2, value)
                    connect.commit()
            # 用户已存在的情况
            else:
                # 查询该用户是否之前曾有过群组的记录
                # sql = "select * from OFW_USER_TAB_ALLOWED where IsDelete = 1 and MemberCode = %s and GroupID= %s"
                sql = "select * from OFW_USER_TAB_ALLOWED_TEST where IsDelete = 1 and MemberCode = %s and GroupID= %s"
                value = (workcode, groupidnew)
                cursor.execute(sql, value)  # 执行sql语句
                rowold = cursor.fetchall()  # 读取查询结果,
                rowCount = len(rowold)
                # 无群组记录
                if rowCount != 0:
                    # sql = "UPDATE OFW_USER_TAB_ALLOWED SET IsDelete =0 WHERE MemberCode = %s and GroupID= %s"
                    sql = "UPDATE OFW_USER_TAB_ALLOWED_TEST SET IsDelete =0 WHERE MemberCode = %s and GroupID= %s"
                    cursor.execute(sql, value)
                    connect.commit()
                    with server.auth.sign_in(tableau_auth):
                        # 将用户加入群组
                        # index = getTwoDimensionListIndex(tbgroup, groupidnew)
                        user = getTwoDimensionListIndex(tbuser, workcode)
                        # print(user)
                        all_groups, pagination_item = server.groups.get(request_options)
                        for group in all_groups:
                            if group.name == groupidnew:
                                server.groups.add_user(group, tbuser[user][0])
                                break
                else:
                    oldgroup = []
                    # 该员工原有的群组
                    for n in row:
                        oldgroup.append(n[1])
                        # api传递的数据是一一对应的情况
                    print(workcode, "该员工对应原有群组", oldgroup)
                    if groupidnew not in oldgroup:
                        with server.auth.sign_in(tableau_auth):
                            # 将用户加入群组
                            # index = getTwoDimensionListIndex(tbgroup, groupidnew)
                            user = getTwoDimensionListIndex(tbuser, workcode)
                            print(user)
                            all_groups, pagination_item = server.groups.get(request_options)
                            for group in all_groups:
                                if group.name == groupidnew:
                                    server.groups.add_user(group, tbuser[user][0])
                                    break

                            # 更新详细信息
                            # newU.name = 'test'
                            # newU.fullname = workcode
                            # sql2 = "insert into OFW_USER_TAB_ALLOWED (MemberCode,GroupID) VALUES (%s,%s)"
                            sql2 = "insert into OFW_USER_TAB_ALLOWED_TEST (MemberCode,GroupID) VALUES (%s,%s)"
                            value = (workcode, groupidnew)
                            cursor.execute(sql2, value)
                            connect.commit()
    f.close()
    cursor = connect.cursor()  # 创建一个游标对象,python里的sql语句都要通过cursor来执行
    # 查询群组表内的群组信息
    # sql = "select GroupName from OFW_GROUP_TAB_ALLOWED"
    sql = "select GroupName from OFW_GROUP_TAB_ALLOWED_TEST"
    cursor.execute(sql)  # 执行sql语句
    rowgroup = cursor.fetchall()  # 读取查询结果,
    # rowCount = len(row)
    for groupid in rowgroup:
        # print(groupid)
        groupid = groupid[0]
        # 查询表内对应该群组的workcode
        # sqls = "select MemberCode from OFW_USER_TAB_ALLOWED where GroupID= %s"
        sqls = "select MemberCode from OFW_USER_TAB_ALLOWED_TEST where GroupID= %s and IsDelete =0"
        cursor.execute(sqls, groupid)  # 执行sql语句
        rowidingroups = cursor.fetchall()  # 读取查询结果,
        # print(rowidingroups)
        for rowidingroup in rowidingroups:
            rawdata = rowidingroup[0]
            # 此处注意判断是否有另外开的账号 需要加在指定群组中 但不在泛微数据中
            if [rawdata, groupid] not in workgroup:
                # 通过tsc 获取id 在group中删除
                # indexgroup=getTwoDimensionListIndex(tbgroup,groupid)
                indexuser = getTwoDimensionListIndex(tbuser, rawdata)
                with server.auth.sign_in(tableau_auth):
                    # mygroup = tbgroup[indexgroup][0]
                    userdel = tbuser[indexuser][0]
                    all_groups, pagination_item = server.groups.get(request_options)
                    for group in all_groups:
                        if group.name == groupid:
                            server.groups.remove_user(group, userdel)
                            # 若该群组不存在 进行逻辑删除
                            # sql = "UPDATE OFW_USER_TAB_ALLOWED SET IsDelete =1 WHERE MemberCode = %s and GroupID= %s"
                            sql = "UPDATE OFW_USER_TAB_ALLOWED_TEST SET IsDelete =1 WHERE MemberCode = %s and GroupID= %s"
                            value = (rawdata, groupid)
                            cursor.execute(sql, value)
                            # 提交到数据库执行
                            connect.commit()
                            break
                    print('完成对 ' + rawdata + " 在群组 " + groupid + " 的删除操作")
    for userid in tbuser:
        # if userid[1] not in SystemUser:
        # 判断是否为非城建用户
        if userid[1] not in SystemUsertest:
            sql = "select * from DW_HR_MEMBER where MemberMsgCode = %s AND snapdate = CONVERT(DATE,GETDATE(),110)"
            cursor.execute(sql, userid[1])  # 执行sql语句
            row = cursor.fetchall()  # 读取查询结果,
            rowCount = len(row)
            # 为0代表用户已离职 从tableau server端删除 计数器+1
            if rowCount == 0:
                with server.auth.sign_in(tableau_auth):
                    server.users.remove(userid[0])
                    print('员工'+ userid[1] +'已离职，从Tableau Server端删除成功')
                    quitCount =quitCount+1

    print('本次共删除用户'+ str(quitCount) +'人')



if __name__ == '__main__':
    main()
    print('End')

