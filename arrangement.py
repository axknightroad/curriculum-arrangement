# -*- coding:utf-8 -*-
########################
##随机排课程序##
##code by axel han##
##2014.12##
######################


#import random


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
class SeasonSchedule:
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
        course.over=start-1+course.clenth
        for i in range(course.clenth):
            self.schedule[weekday][start+i][0]+=1                
            self.schedule[weekday][start+i].append(course.cid)
        course.flag=False
        


class Course:
    'a class for all courses'

    def __init__(self,cnumber,cname,cseason,coptional,cteacher,csubject,clength):
        self.cnumber=cnumber      #课程编号
        self.cname=cname          #课程名称
        self.cseason=cseason      #课程所在学期，0：春，1：夏，2：春夏
        self.coptional=coptional  #课程为选修还是必修, 0：必修，1：选修
        self.cteacher=cteacher    #课程任课老师编号
        self.csubject=csubject    #课程学科点编号
        self.clength=clength      #课程长度
        self.cday=-1              #在星期几上课
        self.cstart=0             #上课开始时间
        self.cover=0              #上课结束时间
        self.flag=True            #课程是否可用，即是否已安排，True:未安排，False：已安排
        self.sub=0                #课程是否经过拆分，0:未经过拆分，1：父课程，2：子课程
        self.priority=0           #课程优先级，越大则优先级越高

    def split(self):  #拆分课程
        child=self
        self.sub=1
        child.sub=2
        return child

            
def selectTime(flag,day,seasonSchedule,course): #为相应课程选择合适时间,flag表示上午0/下午1/晚上2
    length=course.clength
    t=flag*5
    s=1+t
    e=5+t
    for i in range(s,e):
        if seasonSchedule[day][i][0]<maxs and e-s+1>=length:
            ava_flag=True
            for j in range(length):
                if seasonSchedule[day][i+j][0]>=maxs:
                    ava_flag=False
            if ava_flag:
                seasonSchedule.add(day,i,course)
                add_flag=True
                return True
    return False
        
        

def arrangeDay(clist,day,seasonSchedule):  #为一天安排课的函数,flag表示安排课的稠密程度,0:一天只安排一节,1：上午下午最多各一节，2:无限制
    for sno in clist:
        cid=clist[sno]
        course=course_list[cid]
        if course.flag == False:
            continue
        tid=course.cteacher
        if (day,2) in unavailable_dict[tid]:
            continue
        elif (day,0) in unavailable_dict[tid]:
            if not selectTime(1,day,seasonSchedule,course):
                continue
        else:
            if not selectTime(0,day,seasonSchedule,course):
                elif  (day,1)  in unavailable_dict[tid]:                        
                    continue
                elif not selectTime(1,day,seasonSchedule,course):
                    continue
        del clist[sno]
        return True
    return False
            

        
def arrange(clist,seasonSchedule):   #为一类课安排时间的函数
    while clist:
        if not equal_flag:    #假如某一天课程数小于其他日子的课程数，优先
            for i in range(5):
                fail_flag=False
                temp_flag=flag
                while seasonSchedule[i][0]<maxc and not fail_flag:
                    while not arrangeDay(temp_flag,clist,i,seasonSchedule):
                        temp_flag+=1
                        if temp_flag>2:
                            fail_flag=True
                            break
            equal_flag=True
        if not clist:
            return True
        maxc+=1
        add_flag=False    
        for i in range(5):
            if clist:
                temp_flag=flag
                while not arrangeDay(temp_flag,clist,i,seasonSchedule):
                    temp_flag+=1
                    if temp_flag>2:
                        equal_flag=False
                        break
            else:
                equal_flag=False
                break
        if not add_flag:
            maxc-=1
            maxs+=1
            
        
        
    
    
            
        
# for test #
#a=SeasonSchedule()
#a.schedule[0][3].append(31011)
#print a.schedule


#录入数据及预处理#



#随机排课#


#必修课排课
arrange(required_list,0,spring)

#选修课按课程点顺序排课
for optional_list in optional_lists:
    arrange(optional_list,1,spring)
                
            
