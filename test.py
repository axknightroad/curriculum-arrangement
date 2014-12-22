# -*- coding:utf-8 -*-

import xlrd,xlwt,xlutils

weekday_name=['Monday','Tuesday','Wendesday','Thursday','Friday'] #数字对应的星期几
course_list=[]     #课程list
teacher_list=[]    #教师list
subject_list=[]    #学科点list
unavailable_dict={}  #教师时间是否可用dict
required_list=[]   #必修课list
optional_lists=[]   #选修课list，optional_list[i]，表示第i个课程点的选修课list
maxc=0  #单日最大课程数
maxs=1  #一段时间同时上的课程数
equal_flag=True  #是否每天课程数相等
add_flag=False #一轮下来是否增加了课程


class Subject:
    'the subject class'

    def __init__(self,name,start,end):
        self.name=name
        self.start=start
        self.end=end


def get_subject(table):
    print table.ncols,table.nrows
    for i in range(1,table.nrows):
        if table.row(i)[0].value and not table.row(i)[1].value:
            name=table.row(i)[0].value
            i+=1
            start=i
            while i+1<table.nrows and table.row(i+1)[0].value:
                i+=1
            end=i
            subject_list.append(Subject(name,start,end))



def open_excel(file= 'file.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)



data=open_excel('data.xls')
table=data.sheets()[0]
get_subject(table)

for subject in subject_list:
    print subject.name,subject.start,subject.end
