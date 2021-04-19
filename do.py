## -*- coding: utf-8 -*-
import os
import xlrd
#2.0.1版本的xlrd不能操作xlsx文件，将xlrd回退至1.2.0
import xlwt
import win32com.client

def unprotect_xlsx(filename, pw_str):
    xcl = win32com.client.Dispatch("Excel.Application")
    wb = xcl.Workbooks.Open(filename, False, False, None, pw_str)
    xcl.DisplayAlerts = False
    wb.SaveAs(filename, None, '', '')
    xcl.Quit()

outputname='2020年度期货从业人员状况调查汇总表.xls'
if os.path.exists(outputname):
    os.remove(outputname)

#设置输出文件
workbook = xlwt.Workbook(encoding='utf-8')
worksheet=workbook.add_sheet('Sheet1')

#打开配置文件
with open('titles.txt','r') as f:
    l=f.readlines()

#输出表头数据
worksheet.write(0,0,label='序号')
columns=1
for name in l:
    labels=name.split()[0]
    worksheet.write(0,columns,label=labels)
    columns+=1

#设置源excel文件
filepath=os.path.dirname(os.path.abspath(__file__))+"\\"
print(filepath)
filelist = os.listdir(filepath)

rows=1
#遍历目录下所有文件
for xls in filelist:
    if os.path.splitext(xls)[-1]=='.xlsx' or os.path.splitext(xls)[-1]=='.xls':#过滤文件名
        try:
            data = xlrd.open_workbook(filepath+xls)
        except PermissionError:
            print("PermissonError",filepath+xls)
            continue
        except xlrd.biffh.XLRDError:
            print(filepath+xls)
            unprotect_xlsx(filepath+xls,'')
            data = xlrd.open_workbook(filepath+xls)
        
        else:
            #序号
            worksheet.write(rows,0,label=rows)
            table = data.sheets()[0]

            columns=1
            
            for items in l:
                x=int(items.split()[1])-1
                y=ord(items.split()[2])-ord('a')

                is_int=False
                item_list = items.split()
                if len(item_list)==4 and item_list[-1]=='i':
                    is_int=True

                content=table.cell_value(x,y)

                if is_int:
                    try:
                        content=int(content)
                    except ValueError:
                        content=0

                worksheet.write(rows,columns,label=content)
                columns += 1

            rows+=1

workbook.save(outputname)