#!/bin/python
#coding:utf8

import cx_Oracle                                                            #引用模块cx_Oracle
import csv
import sys
import datetime
import time
import xlsxwriter
import openpyxl
from openpyxl import load_workbook
import xlwt
import sys

reload(sys)

sys.setdefaultencoding('gbk')

conn=cx_Oracle.connect('test/test@127.0.0.1/test')         #连接数据库
##excel file
sheet_name =('test')


out_path ='/home/oracle/pythondir/test'+'.xlsx'

cur=conn.cursor()                                                           #获取cursor
query_sql='''select *from dual
'''

re=cur.execute(query_sql)                                                         #使用cursor进行各种操作
rs=cur.fetchall()
#print rs

# 获取MYSQL里面的数据字段名称
fields = cur.description
workbook = xlsxwriter.Workbook(out_path) # workbook是sheet赖以生存的载体。
sheet = workbook.add_worksheet(sheet_name)

# 写上字段信息 
for field in range(0,len(fields)):
    sheet.write(0,field,fields[field][0])

# 获取并写入数据段信息 
row = 1 
col = 0 
for row in range(1,len(rs)+1):
    for col in range(0,len(fields)):
        sheet.write(row,col,u'%s'%rs[row-1][col]) 

#workbook.save(out_path)
cur.close()   
conn.close()
workbook.close()
