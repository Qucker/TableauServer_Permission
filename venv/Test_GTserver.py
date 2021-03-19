import tableauserverclient as TSC
import os
import json
import requests
import pymssql
from urllib import parse
import xml.etree.ElementTree as ET # Contains methods used to build and parse XML

xmlns = {'t': 'http://tableau.com/api'}
# 定义请求header类型
HEADERS = {'Content-Type': 'application/x-www-form-urlencoded;charset=utf-8', 'Key': '332213fa4a9d4288b5668ddd9'}

# 定义请求地址 分别为 公司分部，部门，岗位，员工信息

# urls = [
#      "http://oa.sucgm.cn:8090/api/resource/getHrmSubcompany"
#     ,"http://oa.sucgm.cn:8090/api/resource/getHrmDepartment"
#     ,"http://oa.sucgm.cn:8090/api/resource/getHrmJobtitle"
#     ,"http://oa.sucgm.cn:8090/api/resource/getHrmResource"
# ]
# 通过字典方式定义请求body
# FormData = {"username": 'WBS', "password": '1002AA10000000002BNO', "deptcode": 'WBXT001'}
# 字典转换k1=v1 & k2=v2 模式


def addgroup(groupname):
    # create a new instance with the group name
    # 添加群组
    newgroup = TSC.GroupItem(groupname)

    # call the create method
    newgroup = server.groups.create(newgroup)

    # print the names of the groups on the server
    # all_groups, pagination_item = server.groups.get()
    # for group in all_groups :
    #     print(group.name, group.id)

def _encode_for_display(text):
    """
    Encodes strings so they can display as ASCII in a Windows terminal window.
    This function also encodes strings for processing by xml.etree.ElementTree functions.
    Returns an ASCII-encoded version of the text.
    Unicode characters are converted to ASCII placeholders (for example, "?").
    """
    return text.encode('ascii', errors="backslashreplace").decode('utf-8')


def addusertogroup(username,role,mygroup):
    # create a new UserItem object.
    # 添加用户，分配站点角色
    newU = TSC.UserItem(username,role)

    # add the new user to the site
    # 将用户加入群组
    newU = server.users.add(newU)
    print(newU.name, newU.site_role,newU.id)

    # all_groups, pagination_item = server.groups.get()
    # mygroup = all_groups[1]

    # The id for Ian is '59a8a7b6-be3a-4d2d-1e9e-08a7b6b5b4ba'

    # add Ian to the group
    server.groups.add_user(mygroup, newU.id)


def selectinfo(list,value):
    # i=0
    dict={}
    for each in list:
        if (value in each):
            dict[each.get("outkey")]=each.get(value)
            # i=i+1
    print(dict)
    return dict

def selectHR(list):
    Hrinfo=[]
    i=0
    for each in list:
        dict = {}
        dict["workcode"]=each.get("workcode")
        dict["subcompanyid1"]=each.get("subcompanyid1")
        dict["departmentid"]=each.get("departmentid")
        dict["jobtitle"]=each.get("jobtitle")
        Hrinfo[i]=dict
        i=i+1
    print(Hrinfo)
    return Hrinfo

def select(dict):
    for value in dict:
        groupname=dict.get(value)
        addgroup(groupname)


def inserthr(dict_subcom,dict_depart,dict_job,Hrinfo):
    for hr in Hrinfo:
        hrworkcode=hr.get("workcode")
        hrsubcom=hr.get("subcompanyid1")
        hrdepart=hr.get("departmentid")
        hrjob=hr.get("jobtitle")
        groups=[]
        groups[0]=dict_subcom.get(hrsubcom)
        groups[1]=dict_depart.get(hrdepart)
        groups[2]=dict_job.get(hrjob)
        for group in groups:
            # 需要确定下建立的站点角色
            addusertogroup(hrworkcode,role,group)


def Projectupgrade(groupname):
    # 分配项目权限
    # import tableauserverclient as TSC
    # server = TSC.Server('https://MY-SERVER')
    # sign in, etc

    locked_true = TSC.ProjectItem.ContentPermissions.LockedToProject
    print(locked_true)
    # prints 'LockedToProject'

    by_owner = TSC.ProjectItem.ContentPermissions.ManagedByOwner
    print(by_owner)
    # prints 'ManagedByOwner'

    # pass the content_permissions to new instance of the project item.
    new_project = TSC.ProjectItem(name='My Project', content_permissions=by_owner, description='Project example')


def add_new_permission(server, auth_token, site_id, workbook_id, user_id, permission_name, permission_mode):
    """
    Adds the specified permission to the workbook for the desired user.
    'server'            specified server address
    'auth_token'        authentication token that grants user access to API calls
    'site_id'           ID of the site that the user is signed into
    'workbook_id'       ID of workbook to audit permission in
    'user_id'           ID of the user to audit
    'permission_name'   name of permission to add or update
    'permission_mode'   mode to set the permission
    """
    url = server + "/api/{0}/sites/{1}/workbooks/{2}/permissions".format(3.8, site_id, workbook_id)

    # Build the request
    xml_request = ET.Element('tsRequest')
    permissions_element = ET.SubElement(xml_request, 'permissions')
    ET.SubElement(permissions_element, 'workbook', id=workbook_id)
    grantee_element = ET.SubElement(permissions_element, 'granteeCapabilities')
    ET.SubElement(grantee_element, 'user', id=user_id)
    capabilities_element = ET.SubElement(grantee_element, 'capabilities')
    ET.SubElement(capabilities_element, 'capability', name=permission_name, mode=permission_mode)
    xml_request = ET.tostring(xml_request)

    server_request = requests.put(url, data=xml_request, headers={'x-tableau-auth': auth_token})
    _check_status(server_request, 200)
    print("\tSuccessfully added/updated permission")
    return


class ApiCallError(Exception):
    pass


def _check_status(server_response, success_code):
    """
    Checks the server response for possible errors.
    'server_response'       the response received from the server
    'success_code'          the expected success code for the response
    Throws an ApiCallError exception if the API call fails.
    """
    if server_response.status_code != success_code:
        parsed_response = ET.fromstring(server_response.text)

        # Obtain the 3 xml tags from the response: error, summary, and detail tags
        error_element = parsed_response.find('t:error', namespaces=xmlns)
        summary_element = parsed_response.find('.//t:summary', namespaces=xmlns)
        detail_element = parsed_response.find('.//t:detail', namespaces=xmlns)

        # Retrieve the error code, summary, and detail if the response contains them
        code = error_element.get('code', 'unknown') if error_element is not None else 'unknown code'
        summary = summary_element.text if summary_element is not None else 'unknown summary'
        detail = detail_element.text if detail_element is not None else 'unknown detail'
        error_message = '{0}: {1} - {2}'.format(code, summary, detail)
        raise ApiCallError(error_message)
    return


def get_workbook_id(server, auth_token, user_id, site_id, workbook_name):
    """
    Gets the id of the desired workbook to relocate.
    'server'        specified server address
    'auth_token'    authentication token that grants user access to API calls
    'user_id'       ID of user with access to workbooks
    'site_id'       ID of the site that the user is signed into
    'workbook_name' name of workbook to get ID of
    Returns the workbook id and the project id that contains the workbook.
    """
    url = server + "/api/{0}/sites/{1}/users/{2}/workbooks".format(3.8, site_id, user_id)
    server_response = requests.get(url, headers={'x-tableau-auth': auth_token})
    _check_status(server_response, 200)
    xml_response = ET.fromstring(_encode_for_display(server_response.text))

    # Find all workbooks in the site and look for the desired one
    workbooks = xml_response.findall('.//t:workbook', namespaces=xmlns)
    for workbook in workbooks:
        if workbook.get('name') == workbook_name:
            return workbook.get('id')
    error = "Workbook named '{0}' not found.".format(workbook_name)
    raise LookupError(error)


if __name__ == '__main__':
    # data = parse.urlencode(FormData)
    connect = pymssql.connect('192.168.90.34', 'sa', 'P@ssw0rd', 'Dev_Beta')  # 建立连接
    if connect:
        print("连接成功!")

    cursor = connect.cursor()  # 创建一个游标对象,python里的sql语句都要通过cursor来执行
    names = '220'
    sql = "select * from tab_system_users where name = %s"
    cursor.execute(sql, names)  # 执行sql语句
    row = cursor.fetchone()  # 读取查询结果,
    print(row)
    if row == None:
        sql = "insert into FW_Users () values ()"
        cursor.execute(sql)
    # 登陆tabelau server
#     tableau_auth = TSC.TableauAuth('20040201', 'GT2020!qaz', site_id='test')
    tableau_auth = TSC.TableauAuth('GT_Account', 'GT2020123456', site_id='')
    server = TSC.Server('https://fas-gt.com:8000')
    # server = TSC.Server('https://tab.sucgm.com:10000')
with server.auth.sign_in(tableau_auth):
    # newU = TSC.UserItem('Publishtester','Publisher')
    #     # add the new user to the site
    # newU = server.users.add(newU)
    # newU.name = '111'
    # newU.fullname = 'test'
    # user1 = server.users.update(newU,'P@ssw0rd')
    # # newgroup = TSC.GroupItem('江苏分公司')
    # newgroup2 = TSC.GroupItem('普通员工')
    # # call the create method
    # print(newU.name, newU.site_role, newU.id)
    workbook_id = get_workbook_id('https://fas-gt.com:8000', 'f8Lzj5ZpQgmM8Y-mKTR7vw|iB2PoBy1pMiSGoo0WJQ48RFjTqZuqcFN',  '91acb10f-867e-47b6-9709-0ff4a8b06e45', '82c1a0f8-49fb-4e11-a92c-a20221b2f4b8', 'ZTO_Viz_2019')
    add_new_permission('https://fas-gt.com:8000', 'P37I9FIcS3eiVBd3aOaSOw|vHvfUK0okebhndhruwxMDc13zVlA36g3', '82c1a0f8-49fb-4e11-a92c-a20221b2f4b8', workbook_id, 'd30c6d24-1dff-412c-8caf-df9572d66ccb', 'Read', 'Allow')
    # newgroup = server.groups.create(newgroup)
    # newgroup2 = server.groups.create(newgroup2)
    # server.groups.add_user(newgroup2, newU.id)
    # The id for Ian is '59a8a7b6-be3a-4d2d-1e9e-08a7b6b5b4ba'

    # add Ian to the group
    # server.groups.add_user(mygroup, '59a8a7b6-be3a-4d2d-1e9e-08a7b6b5b4ba')
#         # 建立分部信息群组
#         select(dict_subcom)
#         # 建立部门信息群组
#         select(dict_depart)
#         # 建立岗位信息群组
#         select(dict_job)
#         # 建立各个站点角色
#         inserthr(dict_subcom,dict_depart,dict_job,Hrinfo)
#
#         # 输出group和user
#         all_groups, pagination_item = server.groups.get()
#         mygroup = all_groups[0]
#
#         print(mygroup.name)
#         pagination_item = server.groups.populate_users(mygroup)
#         print(pagination_item)
#         # print the names of the users
#         for user in mygroup.users:
#             print(user.id)
#
#
#             # get the groups on the server
#             all_groups, pagination_item = server.groups.get()
#
#             # print the names of the first 100 groups
#             for group in all_groups:
#                 print(group.name, group.id)
#
#
#
#
# with server.auth.sign_in(tableau_auth):
#     all_groups, pagination_item = server.groups.get()
#
#     mygroup = all_groups[1]
#
#     print(all_groups)
#     print(mygroup)


