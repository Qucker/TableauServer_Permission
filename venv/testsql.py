# -*- coding: utf-8 -*-
"""
Created on Thu Nov 12 14:15:36 2020

@author: hailey.yuan
连sql server传EXCEL数据到数据库，1000行执行一次sql
"""

import xlrd
import pymssql
import datetime

# 连sql server     地址     用户名  密码   数据库
conn = pymssql.connect('10.10.10.40', 'Bao', 'Bao123456', '2020_Verifone')
# 建立cursor
cursor = conn.cursor()
# excel文件
fname = "Oracle - Over-Time Q1-Q4 (for GT China).xlsx"
start_time1 = datetime.datetime.now()
print(start_time1)
# 打开文件
bk = xlrd.open_workbook(fname)
start_time2 = datetime.datetime.now()
print(start_time2)
# 打开工作表
for nums in range(0, len(bk.sheets())):
    sh = bk.sheets()[0]
    # 获取行数
    start_time = datetime.datetime.now()
    sql3 = ''
    start_time3 = datetime.datetime.now()
    print(start_time3)
    # 遍历所有行
    for i in range(1, sh.nrows):
        a = []
        sql = '('
        # 遍历所有列
        for j in range(sh.ncols):
            # 将excel每一列的值用，隔开
            str_sql = str(sh.cell(i, j).value)
            sql += "'" + str_sql.replace("'", "''") + "'" + ','
        # 组合成sql语句(value1,value2,value3,,)
        sql2 = sql.strip(",")
        sql3 += sql2.strip() + '),'
        # 1000行执行一次sql
        if i % 1000 == 0:
            sql3 = sql3.rstrip(",")  # Oracle_vs_Non_Oracle字段只有sheet3有
            sql1 = "insert into OT_Q1_Q4_2020(GL_F_Qtr, SOB_Name, GL_Date, Sales_Division, Sales_Market, Bill_Cust_Num, Bill_Cust_Name, Bill_Location, Bill_Country_Name, Ship_Cust_Num, Ship_Cust_Name, Ship_Location, TrxSource_Sys, Invoice_Source, Rev_Type, Rep_Name, Order_Type, Order_Num, PO_Number, Inv_Num, Item_Accounting_Rule, Price_Uplift_Class_Code, Prod_Code_Desc, Prod_Line_Desc, Prod_Grp_Desc, Prod_Fmly_Desc, Item_Number, Item_Desc, Billing_Qty, Inv_Currency_Code, Unit_Sell_Price_USD, Inv_Amt_T, Inv_Amt_USD, ABSOLUTE_VAL, GL_COGS_USD, REC_CO, REC_Acct, REC_Product, Ship_From_Org, Ship_To_Org, FOB_Code, US_GAAP) values %s " % sql3
            # 执行sql语句
            cursor.execute(sql1)
            sql = ""
            sql3 = ""
    sql3 = sql3.rstrip(",")
    sql1 = "insert into OT_Q1_Q4_2020(GL_F_Qtr, SOB_Name, GL_Date, Sales_Division, Sales_Market, Bill_Cust_Num, Bill_Cust_Name, Bill_Location, Bill_Country_Name, Ship_Cust_Num, Ship_Cust_Name, Ship_Location, TrxSource_Sys, Invoice_Source, Rev_Type, Rep_Name, Order_Type, Order_Num, PO_Number, Inv_Num, Item_Accounting_Rule, Price_Uplift_Class_Code, Prod_Code_Desc, Prod_Line_Desc, Prod_Grp_Desc, Prod_Fmly_Desc, Item_Number, Item_Desc, Billing_Qty, Inv_Currency_Code, Unit_Sell_Price_USD, Inv_Amt_T, Inv_Amt_USD, ABSOLUTE_VAL, GL_COGS_USD, REC_CO, REC_Acct, REC_Product, Ship_From_Org, Ship_To_Org, FOB_Code, US_GAAP) values %s " % sql3
    cursor.execute(sql1)
    # commit提交变更
    conn.commit()
    # 结束时间
    end_time = datetime.datetime.now()
    print(end_time)
    speed = end_time - start_time
    # 打印花费时间
    print(speed)