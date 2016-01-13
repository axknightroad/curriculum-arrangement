# -*- coding:utf-8 -*-
########################
## 随机排课程序 ##
## code by Axel Han ##
## 2014.12 ##
######################


import xlrd,xlwt,xlutils

weekday_name = ['Monday', 'Tuesday', 'Wendesday', 'Thursday', 'Friday'] #数字对应的星期几
course_list = []     #课程list
subject_list = []    #学科点list
required_list = []   #必修课list
maxc = 0  #单日课程数上限
maxs = 1  #同一时间上课数量上限
equal_flag = True  #是否每天课程数相等
add_flag = False #一轮下来是否增加了课程
unicode_dic = {}  #文字对应unicode编码
add_count = 0

NUMBER_COL = 2
SEASON_COL = 1
OPTIONAL_COL = 4
NAME_COL = 3
TEACHER_COL = 11
LENGTH_COL = 7
PRIORITY_COL = 17
LLM_COL = -1
FA_SHUO_COL = 6

SPRING_ROW = 4
SUMMER_ROW = 2
SPRING_SUMMER_ROW = 7
#AUTUMN_ROW=13
#WINTER_ROW=10
#AUTUMN_WINTER_ROW=2
REQUIRE_ROW = 7
SUBJECT_ROW = 1
#LLM_ROW=6

class Subject:  #学科点类
    'the subject class'

    def __init__(self, name, start, end):
        self.name = name   #学科点名称
        self.start = start #学科点课程在xsl文件中的起始行
        self.end = end     #学科点课程在xsl文件中的结束行
        self.optional_list = []  #学科点包含的选修课列表


class Season:  #学期类
    'the courses schedule class'

    def __init__(self):
        self.schedule = [[0] for i in range(5)]  #schedul[0]~schedul[4]分别表示周一到周五，而schedul[i][0]==j表示星期i+1有j节课
        for weekday in self.schedule:    #schedul[i][j]为星期i+1的第j节课的list，其中schedul[i][j][0]==k表示星期i的第j节课同时安排了k节课
            for j in range(14):
                weekday.append([0])

    def add(self, weekday, start, course):  #将课程增加到课程表中，并把课程中的开始星期，开始时间，结束时间置为相应值
        self.schedule[weekday][0] += 1
        course.cday = weekday
        course.start = start
        course.over = start-1+course.clength
        for i in range(course.clength):
            self.schedule[weekday][start+i][0] += 1
            self.schedule[weekday][start+i].append(course_list.index(course))
        course.flag = False



class Course:
    'a class for all courses'

    def __init__(self, cnumber, cname, cseason, coptional, cteacher, csubject, clength, cal, priority, is_llm, fa_shuo):
        self.cnumber = cnumber      #课程编号
        self.cname = cname          #课程名称
        self.cseason = cseason      #课程所在学期，0：春/秋，1：夏/冬，2：春夏/秋冬
        self.coptional = coptional  #课程为选修还是必修, 0：法硕专业学位课，1：科硕专业学位课, 2:专业选修课
        self.cteacher = cteacher    #课程任课老师名称
        self.csubject = csubject    #课程学科点编号
        self.clength = clength      #课程长度
        self.cal = cal              #老师要求的优先级
        self.day = -1               #在星期几上课
        self.start = 0              #上课开始时间
        self.over = 0               #上课结束时间
        self.flag = True            #课程是否可用，即是否已安排，True:未安排，False：已安排
        self.sub = 0                #课程是否经过拆分，0:未经过拆分，1：父课程，2：子课程
        self.priority = priority    #课程优先级，越大则优先级越高
        self.is_llm = is_llm
        self.fa_shuo = fa_shuo
        self.priority += 200 - self.coptional*100
        if self.cseason == 2:
            self.priority += 100
        self.priority += self.clength
        self.copy_flag = False


class DemoCourse:
    'a litele course class for course list'

    def __init__(self, cid, priority):
        self.cid = cid
        self.priority = priority

def select_time(flag, day, season, course, limit_flag, turn): #为相应课程选择合适时间,flag表示上午0/下午1/晚上2,day为星期几,season为春季/夏季,couser为课程
    global maxs, add_flag, add_count
    length = course.clength
    t = flag*5  #用来确定上午,下午和晚上的偏移量
    s = 1+t
    e = 5+t
    for i in range(s,e):
        if season.schedule[day][i][0] < maxs and e-i+1 >= length and course.cal[day][i] >= turn:
            success_flag = True
            for j in range(length):
                if not success_flag:
                    break
                if season.schedule[day][i+j][0] >= maxs or course.cal[day][i+j] < turn:
                    success_flag = False
                    break
                for k in range(season.schedule[day][i+j][0]):
                    temp_id = season.schedule[day][i+j][k+1]
                    if course_list[temp_id].cteacher == course.cteacher:
                        success_flag = False
                        break
                if not course.coptional:
                    for k in range(season.schedule[day][i+j][0]):
                        temp_id = season.schedule[day][i+j][k+1]
                        if not course_list[temp_id].coptional and course.is_llm == course_list[temp_id].is_llm:
                            success_flag = False
                            break
                if limit_flag:
                    for k in range(season.schedule[day][i+j][0]):
                        temp_id = season.schedule[day][i+j][k+1]
                        if course_list[temp_id].csubject == course.csubject and course.is_llm == course_list[temp_id].is_llm and course.fa_shuo == course_list[temp_id].fa_shuo:
                            success_flag = False
                            break
            if success_flag:
                season.add(day, i, course)
                add_flag = True
                add_count += 1
                print "安排了第%d节课" % add_count, course.cname
                return True
    return False



def arrange_day(clist, day, season, limit_flag, turn):  #为一天安排课,clist为待排课list,day为星期几,season为春季/夏季
    for demo_course in clist:
        cid = demo_course.cid
        course = course_list[cid]   #从课程list中获取课程信息,并判断是否该课程已经排过
        am_count = season.schedule[day][1][0]
        pm_count = season.schedule[day][6][0]
        if course.flag == False:
            continue
        ap = int((am_count-pm_count+9)/10)
        if not select_time(ap, day, season, course, limit_flag, turn):
            if not select_time(1-ap, day, season, course, limit_flag, turn):
                continue
        clist.remove(demo_course)
        return True
    return False



def arrange(clist, season):   #为一类课安排时间的函数
    global equal_flag, maxc, maxs, add_flag
    turn = 3 #用来协调老师要求
    fail_count = 0
    limit_flag = True  #是否不再限制课程冲突
    add_maxs = False
    while clist:
        if not equal_flag:  #如果每天个课程不一样,则优先把课排到课少的天数中
            for i in range(5):
                fail_flag = False     #用来判断是否无法把课程添加到该天
                while season.schedule[i][0] < maxc and not fail_flag:
                    if not arrange_day(clist, i, season, limit_flag, turn):
                        fail_flag = True
            equal_flag = True
            for i in range(5):     #判断每天课程是否相同，将equal_flag置位
                if season.schedule[i][0] < maxc:
                    equal_flag = False
        if not clist:
            return True
        maxc += 1
        add_flag = False
        for i in range(5):  #开始从周一到周五分别为每天安排一节课程
            if clist:
                if not arrange_day(clist, i, season, limit_flag, turn):
                    equal_flag = False
            else:
                equal_flag = False
                break
        if not add_flag:   #如果当前同一时间上课数量上限已经无法将所有课排完,则逐渐放宽条件
            maxc -= 1
            fail_count += 1
            if not add_maxs: #如果之前没有增加同一时间课程数上限,则增加
                maxs += 1
                add_maxs = True
            else:
                maxs -= 1
                add_maxs = False
                if turn > 1:    #如果增加同一时间上限数无效,则放宽老师的要求
                    turn -= 1
                elif limit_flag:  #如果还不行,则不再限制课程冲突
                    limit_flag = False
                elif fail_count > 2: #如果放宽条件后还是失败2次以上,则次课排课失败
                    print '\n\n以下课程排课失败\n\n'
                    for demo_course in clist:
                        cid = demo_course.cid
                        print course_list[cid].cname
                    print '\n\n\n'
                    return False
        else:
            add_maxs = False
            fail_count = 0
    return equal_flag


def get_subject(table):   #从xls文件中获取学科点信息并写入subject_list
    #print table.ncols,table.nrows
    for i in range(1, table.nrows):
        if table.row(i)[0].value and not table.row(i)[1].value:
            name = table.row(i)[0].value
            i += 1
            start = i
            while i + 1 < table.nrows and table.row(i+1)[0].value:
                i += 1
            end = i
            subject_list.append(Subject(name, start, end))
            print "subject", start, end

def open_excel(file = 'file.xls'):  #打开excel文件
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception, e:
        print str(e)


def get_unicode(file = 'file.xls'):  #获得中文对应unicode编码
    data = open_excel(file)
    table = data.sheets()[0]
    unicode_dic['Spring'] = table.cell(SPRING_ROW, SEASON_COL).value
    unicode_dic['Summer'] = table.cell(SUMMER_ROW, SEASON_COL).value
    unicode_dic['Spring_Summer'] = table.cell(SPRING_SUMMER_ROW, SEASON_COL).value
    unicode_dic['Require'] = table.cell(REQUIRE_ROW, OPTIONAL_COL).value
    unicode_dic['Fa_shuo'] = table.cell(REQUIRE_ROW, FA_SHUO_COL).value
    if LLM_COL != -1:
        unicode_dic['Llm'] = table.cell(LLM_ROW,LLM_COL).value
        print unicode_dic['Llm']
    print unicode_dic['Spring'], unicode_dic['Summer'], unicode_dic['Spring_Summer'], unicode_dic['Require'], unicode_dic['Fa_shuo']



def input_excel_data(file = 'file.xls'):
    data = open_excel(file)
    table = data.sheets()[0]
    get_subject(table)
    course_number = 0
    for subject in subject_list:
        for i in range(subject.start, subject.end+1):
            if NUMBER_COL != 0:
                cnumber = table.row(i)[NUMBER_COL].value
            else:
                cnumber = course_number
                course_number += 1
            cname = table.row(i)[NAME_COL].value
            if table.row(i)[SEASON_COL].value == unicode_dic['Spring']:  #该课为春季课程
                cseason = 0
            elif table.row(i)[SEASON_COL].value == unicode_dic['Summer']: #该课为夏季课程
                cseason = 1
            elif table.row(i)[SEASON_COL].value == unicode_dic['Spring_Summer']: #该科为春夏课程
                cseason = 2
            else:
                continue
            if table.row(i)[OPTIONAL_COL].value == unicode_dic['Require'] and table.row(i)[FA_SHUO_COL].value == unicode_dic['Fa_shuo']: #该课为必修课
                coptional = 0
            elif table.row(i)[OPTIONAL_COL].value == unicode_dic['Require']:
                coptional = 1
            else:
                coptional = 2
            cteacher = table.row(i)[TEACHER_COL].value
            csubject = subject.name
            if LLM_COL != -1 and table.row(i)[LLM_COL].value == unicode_dic['Llm']:
                    is_llm = True
            else:
                is_llm = False
            print table.row(i)[LENGTH_COL].value,cname
            clength = int(table.row(i)[LENGTH_COL].value)
            if cseason != 2:
                clength *= 2
            cal=[[0] for k in range(5)]
            for weekday in cal:
                for j in range(14):
                    weekday.append(1)
            priority = 10 * int(table.row(i)[PRIORITY_COL].value)
            #print cname
            if table.row(i)[FA_SHUO_COL].value == unicode_dic['Fa_shuo']:
                fa_shuo = True
            else:
                fa_shuo = False
            if int(table.row(i)[PRIORITY_COL].value) == -1:
                print i,cname,"不排"
                continue
            #print int(table.row(i)[PRIORITY_COL].value)
            for k in range(int(table.row(i)[PRIORITY_COL].value)):
                temp_turn = int(table.row(i)[PRIORITY_COL+1+k*4].value)
                temp_day = int(table.row(i)[PRIORITY_COL+1+k*4+1].value)
                #print "temp_day",temp_day
                temp_start = int(table.row(i)[PRIORITY_COL+1+k*4+2].value)
                temp_end = int(table.row(i)[PRIORITY_COL+1+k*4+3].value)
                if temp_day < 5:
                    for j in range(temp_start,temp_end+1):
                        cal[temp_day][j] = temp_turn
                else:
                    for weekday in cal:
                        #print "weekday",weekday
                        for j in range(temp_start, temp_end+1):
                         #   print temp_start,temp_end
                         #   print j
                            weekday[j] = temp_turn
            for l in range(6, 11):
                cal[4][l] = 0
            this_course = Course(cnumber, cname, cseason, coptional, cteacher, csubject, clength, cal, priority, is_llm, fa_shuo)
            course_list.append(this_course)
            print cname
            #if cseason != 2:
            #    copy_course=Course(cnumber+1,cname,cseason,coptional,cteacher,csubject,clength,cal,priority,is_llm)
            #    course_number += 1
            #    course_list.append(copy_course)


def get_course_list(season):
    for course in course_list:
        if course.cseason == season or (course.cseason == 2 and season == 0):
            cid=course_list.index(course)
            priority = course.priority
            this_demo = DemoCourse(cid,priority)
            if course.coptional == 0:
                required_list.append(this_demo)
                #print this_demo.cid
            else:
                for subject in subject_list:
                    if course.csubject == subject.name:
                        subject.optional_list.append(this_demo)
    required_list.sort(lambda p1, p2: cmp(p1.priority, p2.priority), reverse=True)
    for subject in subject_list:
        subject.optional_list.sort(lambda p1,p2 : cmp(p1.priority, p2.priority), reverse=True)



def output_schedule(season, filename, flag):
    file = xlwt.Workbook()
    table = file.add_sheet(u'sheet0',cell_overwrite_ok=True)
    if flag == 0:
        table.write(0, 0, u'autumn\u0020schedule')
    else:
        table.write(0, 0, u'winter\u0020schedule')
    for i in range(5):
        table.write(1, 1+i*10, weekday_name[i])
        for j in range(1, 11):
            table.write(2 + (j-1)*3, 0, u'第%d节课'%j)
            for k in range(season.schedule[i][j][0]):
                cid = season.schedule[i][j][k+1]
                cname = course_list[cid].cname
                table.write(2 + (j-1)*3, 1+ i*10 +k, cname)
    file.save(filename)
    return True

def copy_schedule(spring, summer):
    global maxs, maxc, equal_flag
    for i in range(5):
        for j in range(1, 11):
            for k in range(spring.schedule[i][j][0]):
                cid = spring.schedule[i][j][k+1]
                if course_list[cid].cseason == 2 and not course_list[cid].copy_flag:
                    # print course_list[cid].cname,course_list[cid].cseason
                    start = course_list[cid].start
                    summer.add(i, start, course_list[cid])
                    course_list[cid].copy_flag = True
    maxc = 0
    maxs = 0
    for i in range(5):
        if summer.schedule[i][0] > maxc:
            maxc = summer.schedule[i][0]
        for j in range(1,11):
            if summer.schedule[i][j][0] > maxs:
                maxs = summer.schedule[i][j][0]
    equal_flag = False
    return True



#录入数据及预处理#
get_unicode('2015yaw.xls')
input_excel_data('2015yaw.xls')

spring = Season()
summer = Season()
maxc = 0
maxs = 1



#获取春季及春夏需排课的列表
get_course_list(0)
#for this_course in required_list:
#    print this_course.cid

# 为春季及春夏学季的课程排课 #
arrange(required_list, spring)
for subject in subject_list:
    arrange(subject.optional_list, spring)

# 输出春季排课结果 #
if output_schedule(spring, 'autumn.xls', 0):
    print "autumn successful"

# 清空春季必修课和选修课列表 #
while required_list:
    required_list = []
for subject in subject_list:
    while subject.optional_list:
        subject.optional_list = []

# 将春夏课程的排课结果拷贝到夏季课表中 #
if copy_schedule(spring, summer):
    print "copy successful"

# 获取夏季需排课的列表 #
get_course_list(1)

# 为夏季课程排课 #
arrange(required_list, summer)
for subject in subject_list:
    arrange(subject.optional_list, summer)

# 输出夏季课表 #
if output_schedule(summer, 'winter.xls', 1):
    print "winter successful"
