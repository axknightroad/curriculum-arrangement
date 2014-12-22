# -*- coding:utf-8 -*-
########################
##随机排课程序##
##code by axel han##
##2014.12##
######################


#import random
import xlrd,xlwt,xlutils

weekday_name=['Monday','Tuesday','Wendesday','Thursday','Friday'] #数字对应的星期几
course_list=[]     #课程list
# teacher_list=[]    #教师list
subject_list=[]    #学科点list
unavailable_dict={}  #教师时间是否可用dict
required_list=[]   #必修课list
#optional_lists={}   #选修课list，optional_list[i]，表示第i个课程点的选修课list
maxc = 0  #单日最大课程数
maxs = 1  #一段时间同时上的课程数
equal_flag=True  #是否每天课程数相等
add_flag=False #一轮下来是否增加了课程
chinese=[]

class Subject:
    'the subject class'

    def __init__(self,name,start,end):
        self.name=name
        self.start=start
        self.end=end
        self.optional_list=[]


class Season:
    'the courses schedule class'

    def __init__(self):
        self.schedule=[[0] for i in range(5)]  #schedul[0]~schedul[4]分别表示周一到周五，而schedul[i][0]==j表示星期i+1有j节课
        for weekday in self.schedule:    #schedul[i][j]为星期i+1的第j节课的list，其中schedul[i][j][0]==k表示星期i的第j节课同时安排了k节课
            for j in range(10):
                weekday.append([0])
                
    def add(self,weekday,start,course):  #将课程增加到课程表中，并把课程中的开始星期，开始时间，结束时间置为相应值                    
        self.schedule[weekday][0]+=1
        course.cday=weekday
        course.start=start
        course.over=start-1+course.clength
        for i in range(course.clength):
            self.schedule[weekday][start+i][0]+=1                
            self.schedule[weekday][start+i].append(course_list.index(course))
        course.flag=False
        


class Course:
    'a class for all courses'

    def __init__(self,cnumber,cname,cseason,coptional,cteacher,csubject,clength):
        self.cnumber=cnumber      #课程编号
        self.cname=cname          #课程名称
        self.cseason=cseason      #课程所在学期，0：春，1：夏，2：春夏，3：其它
        self.coptional=coptional  #课程为选修还是必修, 0：必修，1：选修
        self.cteacher=cteacher    #课程任课老师编号
        self.csubject=csubject    #课程学科点编号
        self.clength=clength      #课程长度
        self.day=-1              #在星期几上课
        self.start=0             #上课开始时间
        self.over=0              #上课结束时间
        self.flag=True            #课程是否可用，即是否已安排，True:未安排，False：已安排
        self.sub=0                #课程是否经过拆分，0:未经过拆分，1：父课程，2：子课程
        self.priority=0           #课程优先级，越大则优先级越高
        if self.coptional==0:
            self.priority+=50
        if self.cseason==2:
            self.priority+=100
        self.priority+=self.clength



        
    def split(self):  #拆分课程
        child=self
        self.sub=1
        child.sub=2
        return child


class DemoCourse:
    'a litele course class for course list'

    def __init__(self,cid,priority):
        self.cid=cid
        self.priority=priority
            
def select_time(flag,day,season,course): #为相应课程选择合适时间,flag表示上午0/下午1/晚上2
    global maxs,add_flag
    length=course.clength
    t=flag*5
    s=1+t
    e=5+t
    for i in range(s,e):
        #print i,season.schedule[day][i][0]
        if season.schedule[day][i][0]<maxs and e-s+1>=length:
            ava_flag=True
            for j in range(length):
                if season.schedule[day][i+j][0]>=maxs:
                    ava_flag=False
            if ava_flag:
                season.add(day,i,course)
                add_flag=True
                return True
    return False
        
        

def arrange_day(clist,day,season):  #为一天安排课的函数,flag表示安排课的稠密程度,0:一天只安排一节,1：上午下午最多各一节，2:无限制
    for demo_course in clist:
        cid=demo_course.cid
        course=course_list[cid]
        if course.flag == False:
            continue
        teacher=course.cteacher
        if not unavailable_dict.has_key(teacher) or  not unavailable_dict[teacher][day]:
            if not select_time(0,day,season,course):
                if not select_time(1,day,season,course):
                    continue
        elif unavailable_dict[teacher][day]==1:
            if not select_time(1,day,season,course):
                continue
        elif unavailable_dict[teacher][day]==2:
            if not select_time(0,day,season,course):
                continue
        else:
            continue    
        clist.remove(demo_course)
        return True
    return False
            

        
def arrange(clist,season):   #为一类课安排时间的函数
    global equal_flag,maxc,maxs,add_flag
    while clist:
        if not equal_flag:    #假如某一天课程数小于其他日子的课程数，优先
            for i in range(5):
                fail_flag=False
                while season.schedule[i][0]<maxc and not fail_flag:
                    if not arrange_day(clist,i,season):
                        fail_flag=True
            equal_flag=True
            for i in range(5):
                if season.schedule[i][0]<maxc:
                    equal_flag=False
        if not clist:
            return True
        maxc+=1
        add_flag=False    
        for i in range(5):
            if clist:
                if not arrange_day(clist,i,season):
                    equal_flag=False
            else:
                equal_flag=False
                break
        if not add_flag:
            maxc-=1
            maxs+=1
    return equal_flag
            

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


def get_temp(file='file.xls'):
    data=open_excel(file)
    table=data.sheets()[0]
    chinese.append(table.cell(5,1).value)
    chinese.append(table.cell(9,1).value)
    chinese.append(table.cell(10,1).value)
    chinese.append(table.cell(10,6).value)

        
    
def input_excel_data(file='file.xls'):
    data=open_excel(file)
    table=data.sheets()[0]
    get_subject(table)
    for subject in subject_list:
        for i in range(subject.start,subject.end+1):
            cnumber=table.row(i)[2].value
            cname=table.row(i)[4].value
            # this_season=unicodestring(table.row(i)[1].value)
            if table.row(i)[1].value==chinese[0]:
                #print "春"
                cseason=0
            elif table.row(i)[1].value==chinese[1]:
                #print "夏"
                cseason=1
            elif table.row(i)[1].value==chinese[2]:
                #print "春夏"
                cseason=2
            else:
                #print "其它"
                cseason=3
            if table.row(i)[6].value==chinese[3]:
               # print '必修'
                coptional=0
            else:
               # print '选修'
                coptional=1
            cteacher=table.row(i)[12].value
            csubject=subject.name
            if table.row(i)[8]>=2:
                clength=int(table.row(i)[8].value)
            else:
                clength=2
            this_course=Course(cnumber,cname,cseason,coptional,cteacher,csubject,clength)
            course_list.append(this_course)

def get_course_list(season):
    for course in course_list:
        if course.cseason==season or ( course.cseason==2 and season==0):
            cid=course_list.index(course)
            priority=course.priority
            this_demo=DemoCourse(cid,priority)
            if course.coptional == 0:
                required_list.append(this_demo)
            else:
                for subject in subject_list:
                    if course.csubject == subject.name:
                        subject.optional_list.append(this_demo)
    required_list.sort(lambda p1,p2:cmp(p1.priority,p2.priority),reverse=True)
    for subject in subject_list:
        subject.optional_list.sort(lambda p1,p2:cmp(p1.priority,p2.priority),reverse=True)



def output_schedule(season):
    file=xlwt.Workbook()
    table=file.add_sheet(u'sheet0',cell_overwrite_ok=True)
    for i in range(5):
        table.write(0,1+i*5,weekday_name[i])
        for j in range(1,11):
            table.write(1+(j-1)*3,0,u'第%d节课'%j)
            for k in range(season.schedule[i][j][0]):
                cid=spring.schedule[i][j][k+1]
                cname=course_list[cid].cname
                table.write(1+(j-1)*3,1+i*5+k,cname)
    file.save('out.xls')
    return True            
    
        
# for test #
#a=SeasonSchedule()
#a.schedule[0][3].append(31011)
#print a.schedule


#录入数据及预处理#
get_temp('data.xls')
input_excel_data('data.xls')
get_course_list(0)
spring=Season()
summer=Season()
maxc=0
maxs=1
#for course in required_list:
#    cid=course.cid
#    print "课程名称: ",course_list[cid].cname,"优先级: ",course.priority,"学季：",course_list[cid].cseason


#随机排课#


#必修课排课
arrange(required_list,spring)
for subject in subject_list:
    arrange(subject.optional_list,spring)

if output_schedule(spring):
    print "successful"
#for i in range(5):
#    print weekday_name[i]
#    for j in range(1,10):
#        print "第%d节课"%j
#        for k in range(spring.schedule[i][j][0]):
#            cid=spring.schedule[i][j][k+1]
#            print course_list[cid].cname
#        print
#        print
        


#选修课按课程点顺序排课
#for optional_list in optional_lists:
#    arrange(optional_list,1,spring)
                

#输出结果
# print chinese

# print chinese


#course_list.sort(lambda p1,p2:cmp(p1.priority,p2.priority),reverse=True)


#for course in course_list:
#    print course.cnumber,course.cname,course.cseason,course.coptional,course.cteacher,course.csubject,course.clength,course.priority
