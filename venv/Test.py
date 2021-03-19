import tableauserverclient as TSC
import os
tableau_auth = TSC.TableauAuth('GT_Test', 'Gtsh20200221', site_id='test')
server = TSC.Server('http://43.254.222.13:10000')
# found = [proj for proj in TSC.Pager(server.projects) if proj.name == 'Testing']
#
# print(found[0].id)

def dst_dir_file(file_path, file_ext=".twb", list=[]):

   for item in os.listdir(file_path):
      path = os.path.join(file_path, item)
      if os.path.isdir(path):
         dst_dir_file(path, file_ext, list)
      if path.endswith(file_ext):
         list.append(path)
   return list

with server.auth.sign_in(tableau_auth):
   # create a workbook item
   # wb_item = TSC.WorkbookItem(name='Sample', project_id='1f2f3e4e-5d6d-7c8c-9b0b-1a2a3f4f5e6e')
   # call the publish method with the workbook item
   wb_item = TSC.WorkbookItem('425e367f-7652-4baa-bf36-86517045bb28')
   # 获取路径下所有/指定tableau文件
   list=dst_dir_file('D:\BaiduNetdiskDownload\工作\城建市政\人力资源')
   print(list)
   for item in list:
   # wb_item = server.workbooks.publish(wb_item, 'D:\BaiduNetdiskDownload\工作\城建市政\人力资源\拟退休人员名单.twbx', 'CreateNew')
      wb_item = server.workbooks.publish(wb_item, item, 'CreateNew')

# with server.auth.sign_in(tableau_auth):
#   all_workbooks_items, pagination_item = server.workbooks.get()
#   # print names of first 100 workbooks
#   print([workbook.name for workbook in all_workbooks_items])