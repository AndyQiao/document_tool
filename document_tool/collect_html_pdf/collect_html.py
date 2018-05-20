#! /usr/bin/env python   
# -*- coding: utf-8 -*-

import codecs
import os
import xlrd
import xlwt

    
def CollectHtml(readFile):
    print('start to deal:', readFile)
    res = []
    content = GetHtml(readFile)
    res.append(FindTarget(content, '战略理解'))
    res.append(FindTarget(content, '战略执行'))
    res.append(FindTarget(content, '工作安排'))
    res.append(FindTarget(content, '轻重缓急'))

    res.append(FindTarget(content, '合理分配'))
    res.append(FindTarget(content, '交待仸务'))
    res.append(FindTarget(content, '监督检查'))
    res.append(FindTarget(content, '仸务监控'))

    res.append(FindTarget(content, '反馈技巧'))
    res.append(FindTarget(content, '考核方法'))
    res.append(FindTarget(content, '绩效沟通'))
    res.append(FindTarget(content, '绩效指标'))

    res.append(FindTarget(content, '团队搭配'))
    res.append(FindTarget(content, '团队氛围'))
    res.append(FindTarget(content, '冲突解决'))
    res.append(FindTarget(content, '仸务指导'))

    res.append(FindTarget(content, '个人収展'))
    res.append(FindTarget(content, '沟通能力'))
    res.append(FindTarget(content, '协调能力'))
    return res


def GetHtml(readFile):
    content = ""
    f = codecs.open(readFile,'r','utf-8')
    #f = open(readFile, 'r') 
    lines= f.readlines()
    #lines=[line.decode('utf-8') for line in f.readlines()]
    f.close()
    for i in lines:
        content += i
    return content


def GetLevel(target):
    if(-1 != target.find('高')):
        return '高'
    if(-1 != target.find('中')):
        return '中'
    if(-1 != target.find('低')):
        return '低'
    return ''

def FindTarget(content, object):
    while(1):
        index = content.find(object)
        if(-1 == index):
            return
        target = content[index:100+index]
        level = GetLevel(target)
        if('' != level):
            return level
        content = content[index+20:]
    return '-'






#mystr=os.popen("pdf2htmlEX.exe 1.pdf")





#写入
def Write2Excel(names, levels):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('My Worksheet')
    for row in range(0, len(levels)):
        worksheet.write(row+1, 0, names[row])
        for col in range(0,len(levels[row])):
            worksheet.write(row+1, col+1, levels[row][col])
    workbook.save('out.xls')



rootdif = 'reports_summary'
lists = os.listdir(rootdif)

#名字
names = []
levels =  []
for list in lists:
    #名字
    name = list[0: list.find('-')]
    names.append(name)

    #level
    #level = ["","","","", "","","","", "","","","", "","","","", "","",""]
    levels.append(CollectHtml('reports_summary/' + list))

Write2Excel(names,levels)