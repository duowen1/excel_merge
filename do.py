## -*- coding: utf-8 -*-
import os
import xlrd
import xlwt

titles=['序号','公司名称','问卷填写人','办公电话','手机','电子邮箱',
'责任人','任职部门','职务','联系方式',
'证券从业资格（人）','基金从业资格（人）','律师从业资格（人）','注册会计师（人）','特许金融分析师（人）','金融风险管理师（人）',
'海外留学经历（人）','海外工作经历（人）','海外留学&工作（人）',
'交易所培训（人天）','公司内部培训（人天）','其他培训（人天）','培训总频次（人天）','费用总支出',
'校园招聘（人）','社会招聘-期货公司（人）','社会招聘-基金（人）','社会招聘-金融机构（人）','社会招聘-现货行业（人）','社会招聘-其他（人）','社会招聘（人）',
'离职人员（人）','离职-期货（人）','离职-基金（人）','离职-金融（人）','离职-现货（人）','离职-其他（人）','离职原因-首要','离职原因-次要','离职原因-第三','离职原因-其他',
'资产管理子公司','人员数量','较2018年末增加','风险管理子公司','人员数量','较2018年年末增加','子公司人员数量变化原因','期货行业吸引力得分','意见建议','期货从业人员管理得分','意见建议']



workbook = xlwt.Workbook(encoding='utf-8')
worksheet=workbook.add_sheet('Sheet1')

colums=0
for labels in titles:
    worksheet.write(0,colums,label=labels)
    colums+=1

filepath='.\\query\\'
filelist = os.listdir(filepath)
for xls in filelist:
    if os.path.splitext(xls)[-1]=='.xls':
        print(xls)
        data = xlrd.open_workbook(filepath+xls)
        table = data.sheets()[0]
        



workbook.save('2020年度期货从业人员状况调查汇总表.xls')