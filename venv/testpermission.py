import requests
from xml.dom import minidom #导入这个模块
from xml.dom.minidom import parseString
# capability-name	The capability to assign. Valid capabilities for a workbook are
# AddComment, ChangeHierarchy, ChangePermissions, Delete, ExportData, ExportImage,
# ExportXml, Filter, Read (view), ShareView, ViewComments, ViewUnderlyingData, WebAuthoring, and Write.
# Valid capabilities for a data source are ChangePermissions, Connect, Delete, ExportXml, Read (view), and Write.


def signin():
    url = "https://fas-gt.com:8000/api/3.8/auth/signin"

    payload = "<tsRequest> <credentials name=\"GT_Account\" password=\"GT2020123456\" >\r\n    \t<site contentUrl=\"\" />\r\n\t</credentials>\r\n</tsRequest>"
    headers = {
      'Content-Type': 'text/plain'
    }

    response = requests.request("POST", url, headers=headers, data = payload)
    dom = response.text


    xml_dom = parseString(dom)
    print(dom)
    root =  xml_dom.documentElement
    token = root.getAttribute("credentials")
    print(token)  # 打印节点信息
    print(root.nodeValue)

    print(root.nodeType)
    print(response.text.encode('utf8'))


def permissionadd():
    url = "https://fas-gt.com:8000/api/3.8/sites/82c1a0f8-49fb-4e11-a92c-a20221b2f4b8/workbooks/e29d4d70-3154-40df-a56a-2cb822995ec7/permissions"

    payload = "<tsRequest>\r\n    <permissions>\r\n        <workbook id=\"e29d4d70-3154-40df-a56a-2cb822995ec7\" />\r\n        <granteeCapabilities>\r\n            <user id=\"91acb10f-867e-47b6-9709-0ff4a8b06e45\" />\r\n        </granteeCapabilities>\r\n        <granteeCapabilities>\r\n            <group id=\"b684215b-c085-44bd-b3a1-efcc6bdaeadf\" />\r\n            <capabilities>\r\n                <capability name=\"Filter\" mode=\"Allow\" />\r\n            </capabilities>\r\n        </granteeCapabilities>\r\n\r\n    </permissions>\r\n</tsRequest>\r\n"
    headers = {
        'X-Tableau-Auth': 'iZAG_501QFyK5JaglWdjqA|91fTauPl38qTbmh2HIPQaOjZVrwHgJZ7',
        'Content-Type': 'application/xml'
    }

    response = requests.request("PUT", url, headers=headers, data=payload)

    print(response.text.encode('utf8'))

if __name__ == '__main__':
    signin()