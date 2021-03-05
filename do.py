## -*- coding: utf-8 -*-
import os
import xlrd
#2.0.1版本的xlrd不能操作xlsx文件，将xlrd回退至1.2.0
import xlwt

titles=['序号','公司名称','问卷填写人','办公电话','手机','电子邮箱',
'责任人','任职部门','职务','联系方式',
'证券从业资格（人）','基金从业资格（人）','律师从业资格（人）','注册会计师（人）','特许金融分析师（人）','金融风险管理师（人）',
'海外留学经历（人）','海外工作经历（人）','海外留学&工作（人）',
'交易所培训（人天）','公司内部培训（人天）','其他培训（人天）','培训总频次（人天）','费用总支出',
'校园招聘（人）','社会招聘-期货公司（人）','社会招聘-基金（人）','社会招聘-金融机构（人）','社会招聘-现货行业（人）','社会招聘-其他（人）','社会招聘（人）',
'离职人员（人）','离职-期货（人）','离职-基金（人）','离职-金融（人）','离职-现货（人）','离职-其他（人）','离职原因-首要','离职原因-次要','离职原因-第三','离职原因-其他',
'资产管理子公司','人员数量','较2018年末增加','风险管理子公司','人员数量','较2018年年末增加','子公司人员数量变化原因','期货行业吸引力得分','意见建议','期货从业人员管理得分','意见建议']

def add_content(row,column,l,table,flag=False):
    s=table.cell_value(row,column)
    if not flag:
        l.append(s)
    else:
        try:
            l.append(int(s))
        except ValueError:
            l.append(0)

workbook = xlwt.Workbook(encoding='utf-8')
worksheet=workbook.add_sheet('Sheet1')

colums=0
for labels in titles:#输出表头数据
    worksheet.write(0,colums,label=labels)
    colums+=1

filepath='.\\'
filelist = os.listdir(filepath)
rows=1
for xls in filelist:#遍历目录下所有文件
    if os.path.splitext(xls)[-1]=='.xlsx' or os.path.splitext(xls)[-1]=='.xls':#过滤文件名
        try:
            data = xlrd.open_workbook(filepath+xls)
        except PermissionError:
            continue
        else:
            worksheet.write(rows,0,label=rows)
            table = data.sheets()[0]
            
            content=[]

            #期货公司基本信息
            add_content(3,2,content,table)
            add_content(5,1,content,table)
            add_content(5,3,content,table,True)
            add_content(5,5,content,table,True)
            add_content(5,7,content,table)
            add_content(6,1,content,table)
            add_content(6,3,content,table)
            add_content(6,5,content,table)
            add_content(6,7,content,table,True)

            #期货从业人员结构信息
            ##专业化人才储备情况
            for p in range(9,15):
                add_content(p,7,content,table,True)

            ##国际化人才储备情况
            for p in range(16,19):
                add_content(p,7,content,table,True)

            #就业培训情况
            for p in range(20,25):
                add_content(p,7,content,table)
            
            #从业人员流动情况
            for p in range(26,39):
                add_content(p,7,content,table,True)

            for p in range(39,43):
                add_content(p,6,content,table)

            #期货公司子公司情况
            add_content(44,1,content,table)
            add_content(44,4,content,table,flag=True)
            add_content(44,7,content,table,flag=True)
            add_content(45,1,content,table)
            add_content(45,4,content,table,flag=True)
            add_content(45,7,content,table,flag=True)
            add_content(46,1,content,table)

            #人才队伍建设意见建议
            add_content(48,7,content,table,flag=True)
            add_content(49,0,content,table)
            add_content(50,7,content,table,flag=True)
            add_content(51,0,content,table)

            col=1
            for con in content:
                worksheet.write(rows,col,label=con)
                col+=1

            rows+=1
        
workbook.save('2020年度期货从业人员状况调查汇总表.xls')