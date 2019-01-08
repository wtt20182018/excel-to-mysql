
#excel-to-mysql.py
#encoding=utf-8

import xlrd
from configparser import ConfigParser
import pymysql
import sys

try:
	book = xlrd.open_workbook("xxxxxx.xlsx")  #filename=xxxxxx.xlsx, in the same dir with excel-to-mysql.py
except:
	print("open excel file failed!")
try:
	sheet = book.sheet_by_name("worksheet1")   #worksheet1 is the worksheet in xxxxxx.xlsx
except:
	print("locate worksheet in excel failed!")


#connect database
try:
	database = pymysql.connect(host="XXXXXXXXXX.mysql.rds.aliyuncs.com",user="xxxxxx_rw",
        passwd="xxxxxxxxxx",
        db="xxxxxxxxx",
        charset='utf8')
except:
	print("could not connect to mysql server(XXXXXXXXXX.mysql.rds.aliyuncs.com)")
cursor = database.cursor()
select = "select count(id) from xxxxx" #get item xxxxx from execel 
cursor.execute(select) #excute sql
line_count = cursor.fetchone()
#print(line_count[0])

for i in range(1, sheet.nrows): #the firt line in excel is headline, which is conrespond with database table's field name. So I get data from second line (line 1) in excel. 

	intention_type = sheet.cell(i,0).value #get data in line i  row 0
	iphone = sheet.cell(i,1).value #get data in line i  row 1
  #the rest can be done in the same manner
	addrs = sheet.cell(i,2).value
	intention_date = sheet.cell(i,3).value
	uname = sheet.cell(i,4).value
	ID_card = sheet.cell(i,5).value
	intention_car = sheet.cell(i,6).value
	comment = sheet.cell(i,7).value
#	create_time = now()
	value = (uname,ID_card,iphone,addrs,intention_car,intention_type,intention_date,comment)
	insert = "INSERT INTO dm_intention_pool(name,id_card_no,mobile,addr,model_code,type,intention_time,ext)VALUES(%s,%s,%s,%s,%s,%s,%s,%s)"
	cursor.execute(insert,value) #excute sql


update = "UPDATE `xxxxx`  SET `create_time` =now(),`update_time` =now() where id > %s"
cursor.execute(update,line_count[0]) #excute sql
database.commit()#submit

cursor.close() #close connection
database.close()#close database
print ("")
print ("Done! ")
print ("")

#excel-to-mysql.py end
