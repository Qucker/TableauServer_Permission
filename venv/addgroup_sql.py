import pymssql
import requests
from urllib import parse
import json
import tableauserverclient as TSC


# 连接数据库
def sqlcon():
    connect = pymssql.connect('192.168.90.34', 'sa', 'P@ssw0rd', 'Dev_Beta')  # 建立连接
    if connect:
        print("连接成功!")
        return connect


# 读取泛微Json数据
def fwapi():
    # 建立测试环境连接
    # FormData = {"username": 'WBS', "password": '1002AA10000000002BNO', "deptcode": 'WBXT001'}
    # url = 'http://oa.sucgm.cn:8090/api/resource/getHrmRoles'
    # 建立生产环境连接
    FormData = {"username": 'WBS', "password": '1002AA10000000002BNO', "deptcode": 'WBXT001'}
    url = 'https://cp.sucgm.com/api/resource/getHrmRoles'
    data = parse.urlencode(FormData)
    HEADERS = {'Content-Type': 'application/x-www-form-urlencoded;charset=utf-8'}
    content = requests.post(url=url, headers=HEADERS, data=data).text
    content = json.loads(content)
    return content


def main():
    content = fwapi()
    # tableau_auth = TSC.TableauAuth('20040201', 'GT2020!qaz', site_id='')
    tableau_auth = TSC.TableauAuth('GT_Account', 'GT2020123456', site_id='')
    server = TSC.Server('https://fas-gt.com:8000')
    # server = TSC.Server('https://tab.sucgm.com:10000')
    infos = content['data']
    print(infos)
    connect = sqlcon()
    cursor = connect.cursor()  # 创建一个游标对象,python里的sql语句都要通过cursor来执行
    allgroup = []

    for info in infos:
        # 遍历获取群组id和群组名称
        groupid = info['id']
        groupname = info['rolesmark']  # 获取群组名称
        groupname = groupname[:-5] # 将群组名称后的 -致同群组 字段删除
        allgroup.append([groupid, groupname])

        sql = "select GroupName from OFW_GROUP_TAB_ALLOWED where IsDelete = 0 and GroupId = %s"
        cursor.execute(sql, groupid)
        row = cursor.fetchall()  # 读取查询结果,
        # print(row)
        rowCount = len(row)
        if rowCount == 0:
            sql = "insert into OFW_GROUP_TAB_ALLOWED (GroupId,GroupName) VALUES (%s,%s)"
            value = (groupid, groupname)
            cursor.execute(sql, value)
            # 提交到数据库执行
            connect.commit()
            print('群组' + groupname + '添加成功')
        else:
            rawdata = row[0][0]
            if rawdata != groupname:
                print('群组' + rawdata + '名称需要修改')
                sql = "Update OFW_GROUP_TAB_ALLOWED SET GroupName = %s where GroupId =%s"
                value = (groupname, groupid)
                cursor.execute(sql, value)
                # 提交到数据库执行
                connect.commit()
                print("完成对 " + groupname + " 群组名称的更改")
            else:
                print('已存在' + groupname + '，无需更改')

if __name__ == '__main__':
    main()
