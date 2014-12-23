# -*- coding:utf-8 -*-
########################
##随机排课程序##
##code by axel han##
##2014.12##
######################


import xlrd,xlwt,xlutils

weekday_name=['Monday','Tuesday','Wendesday','Thursday','Friday'] #数字对应的星期几
course_list=[]     #课程list
subject_list=[]    #学科点list
unavailable_dict={}  #教师时间是否可用dict
required_list=[]   #必修课list
maxc=0  #单日课程数上限
maxs=1  #同一时间上课数量上限
equal_flag=True  #是否每天课程数相等
add_flag=False #一轮下来是否增加了课程
unicode_list=[]  #文字对应unicode编码

class Subject:  #学科点类
    'the subject class'

    def __init__(self,name,start,end):
        self.name=name   #学科点名称
        self.start=start #学科点课程在xsl文件中的起始行
        self.end=end     #学科点课程在xsl文件中的结束行
        self.optional_list=[]  #学科点包含的选修课列表


class Season:  #学期类
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
        self.copy_flag=False    



        
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
            
def select_time(flag,day,season,course): #为相应课程选择合适时间,flag表示上午0/下午1/晚上2,day为星期几,season为春季/夏季,couser为课程
    global maxs,add_flag
    length=course.clength
    t=flag*5  #用来确定上午,下午和晚上的偏移量
    s=1+t
    e=5+t 
    for i in range(s,e):
        if season.schedule[day][i][0]<maxs and e-i+1>=length:
            success_flag=True
            for j in range(length):
                if season.schedule[day][i+j][0]>=maxs:
                    success_flag=False
            if success_flag:
                season.add(day,i,course)
                add_flag=True
                return True
    return False
        
        

def arrange_day(clist,day,season):  #为一天安排课,clist为待排课list,day为星期几,season为春季/夏季
    for demo_course in clist:
        cid=demo_course.cid
        course=course_list[cid]   #从课程list中获取课程信息,并判断是否该课程已经排过
        if course.flag == False:    
            continue
        teacher=course.cteacher
        if not unavailable_dict.has_key(teacher) or  not unavailable_dict[teacher][day]: #根据教师有空时间安排课程
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
        if not equal_flag:  #如果每天个课程不一样,则优先把课排到课少的天数中
            for i in range(5):
                fail_flag=False     #用来判断是否无法把课程添加到该天
                while season.schedule[i][0]<maxc and not fail_flag:
                    if not arrange_day(clist,i,season):
                        fail_flag=True
            equal_flag=True
            for i in range(5):     #判断每天课程是否相同，将equal_flag置位
                if season.schedule[i][0]<maxc:
                    equal_flag=False
        if not clist:
            return True
        maxc+=1
        add_flag=False    
        for i in range(5):  #开始从周一到周五分别为每天安排一节课程
            if clist:
                if not arrange_day(clist,i,season):
                    equal_flag=False
            else:
                equal_flag=False
                break
        if not add_flag:   #如果当前同一时间上课数量上限已经无法将所有课排完,则增加当前同一时间上课数量上限
            maxc-=1
            maxs+=1
    return equal_flag
            

def get_subject(table):   #从xls文件中获取学科点信息并写入subject_list
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
            
def open_excel(file= 'file.xls'):  #打开excel文件
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)


def get_unicode(file='file.xls'):  #获得中文对应unicode编码
    data=open_excel(file)
    table=data.sheets()[0]
    unicode_list.append(table.cell(5,1).value)
    unicode_list.append(table.cell(9,1).value)
    unicode_list.append(table.cell(10,1).value)
    unicode_list.append(table.cell(10,6).value)

        
    
def input_excel_data(file='file.xls'):
    data=open_excel(file)
    table=data.sheets()[0]
    get_subject(table)
    for subject in subject_list:
        for i in range(subject.start,subject.end+1):
            cnumber=table.row(i)[2].value
            cname=table.row(i)[4].value
            if table.row(i)[1].value==unicode_list[0]:  #该课为春季课程
                cseason=0
            elif table.row(i)[1].value==unicode_list[1]: #该课为夏季课程
                cseason=1
            elif table.row(i)[1].value==unicode_list[2]: #该科为春夏课程
                cseason=2
            else:         
                continue
            if table.row(i)[6].value==unicode_list[3]: #该课为必修课
                coptional=0
            else:  #该课为选修课
                coptional=1
            cteacher=table.row(i)[12].value
            csubject=subject.name
            clength=int(table.row(i)[8].value)
            if cseason!=2:
                clength*=2
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



def output_schedule(season,filename,flag):
    file=xlwt.Workbook()
    table=file.add_sheet(u'sheet0',cell_overwrite_ok=True)
    if flag==0:
        table.write(0,0,u'spring\u0020schedule')
    else:
        table.write(0,0,u'summer\u0020schedule')
    for i in range(5):
        table.write(1,1+i*10,weekday_name[i])
        for j in range(1,11):
            table.write(2+(j-1)*3,0,u'第%d节课'%j)
            for k in range(season.schedule[i][j][0]):
                cid=season.schedule[i][j][k+1]
                cname=course_list[cid].cname
                table.write(2+(j-1)*3,1+i*10+k,cname)
    file.save(filename)
    return True            

def copy_schedule(spring,summer):
    global maxs,maxc,equal_flag
    for i in range(5):
        for j in range(1,11):
            for k in range(spring.schedule[i][j][0]):
                cid=spring.schedule[i][j][k+1]
                if course_list[cid].cseason == 2 and not course_list[cid].copy_flag:
                    # print course_list[cid].cname,course_list[cid].cseason
                    start=course_list[cid].start
                    summer.add(i,start,course_list[cid])
                    course_list[cid].copy_flag=True
    maxc=0
    maxs=0                
    for i in range(5):
        if summer.schedule[i][0]>maxc:
            maxc=summer.schedule[i][0]
        for j in range(1,11):
            if summer.schedule[i][j][0]>maxs:
                maxs=summer.schedule[i][j][0]
    equal_flag=False
    return True            
    


#录入数据及预处理#
get_unicode('data.xls')
input_excel_data('data.xls')

spring=Season()
summer=Season()
maxc=0
maxs=1


#获取春季及春夏需排课的列表
get_course_list(0)


# 为春季及春夏学季的课程排课 #
arrange(required_list,spring)
for subject in subject_list:
    arrange(subject.optional_list,spring)

# 输出春季排课结果 #
if output_schedule(spring,'spring.xls',0):
    print "spring successful"

# 清空春季必修课和选修课列表 #    
while required_list:
    required_list=[]    
for subject in subject_list:    
    while subject.optional_list:
        subject.optional_list=[]

# 将春夏课程的排课结果拷贝到夏季课表中 # 
if copy_schedule(spring,summer):
    print "copy successful"

# 获取夏季需排课的列表 #
get_course_list(1)

# 为夏季课程排课 #
arrange(required_list,summer)
for subject in subject_list:
    arrange(subject.optional_list,summer)

# 输出夏季课表 #
if output_schedule(summer,'summer.xls',1):
    print "summer successful"
