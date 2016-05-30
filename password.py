import win32clipboard as w
import win32con
import xlrd
from xlrd import open_workbook
from xlutils.copy import copy
import sys 
import os
#set clip board
def setText(aString):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_TEXT, aString)
    w.CloseClipboard()
directory=os.getcwd() 
pass_dir=str(directory+'\password.xls') #directory of password.xls
data=xlrd.open_workbook(pass_dir)
table=data.sheet_by_index(0)
sid=table.col_values(2)
password_all=table.col_values(4)
username_all=table.col_values(3)
col_num=len(sid)
input_sid=raw_input('which password do you need? or show all? ')
input_sid=input_sid.decode(sys.stdin.encoding)

if input_sid=='show all':
	for i in range(col_num):
		print sid[i],username_all[i],password_all[i]
	raw_input()
		
else:
	i=0
	while i<=col_num:
		if input_sid==sid[i]:
			password=str(table.cell(i,4).value)
			username=str(table.cell(i,3).value)
			print "username:",username
			print	"password:",password
			print "Password is already in your clipboard"
			setText(password)
			change=raw_input("Do you want to change it? y or n:")
			if change=="y":
				new_password=raw_input("new password:")
				rb = open_workbook('H:\\grundfos_doc\\administration\\password.xls')
				rs = rb.sheet_by_index(0)
				wb = copy(rb)
				ws = wb.get_sheet(0)
				ws.write(i, 4, new_password)
				wb.save('H:\\grundfos_doc\\administration\\password.xls')
				print "password changed!"
				raw_input()
				break
			else:
				break
		else:
			i=i+1

    
