# -*- coding:utf-8 -*-
########################
## 随机排课程序改 ##
## code by Axel Han ##
## 2015.12 ##
######################

import xlrd,xlwt,xlutils

weekday_name = ['Monday', 'Tuesday', 'Wendesday', 'Thursday', 'Friday'] #数字对应的星期几
unicode_dic = {}  #文字对应unicode编码

#以下为excel中列常量，列号在excel中从1开始，而在程序中应当
NUMBER_COL = ord('c') - ord('a') #表示课程编号在excle文件中的列号
NAME_COL = 4 - 1  #表示课程名字在excle中的列号
LENGTH_COL = 7 - 1  #表示课程长度在excle中的列号
SEASON_COL = 2 - 1  #表示课程所属学期在excle中的列号
YEAR_COL = 5 - 1   #表示课程属性所在列号，一般是上课的年级，若该列为2015表示该课程时2015级上的课,一般选修课为0，必修课为1
PRIORITY_COL = 16 - 1  #表示课程安排要求数量所在的列
OPTIONAL_COL = 6 - 1  #表示课程是否是选修课所在的列


SPRING_ROW = 3 - 1  #某个春学期课程所在行
SUMMER_ROW = 9 - 1  #某个夏学期课程所在行
SPRING_SUMMER_ROW = 0 - 1  #某个春夏学期课程所在行
REQUIRE_ROW = 3 - 1  #某个学位课所在行

klineStart = 3 - 1  #课程开始的行
klineEnd = 12 - 1  #课程结束的行

def open_excel(file = 'file.xls'):  #打开excel文件
    """ 打开excel文件的函数

        如果打开正确则返回该表格数据
        打开错误则返回错误信息
    """
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception, e:
        print str(e)


class Period(object):
    """ 课表中每天的每节课的类

    Attributes:
        isUsed: 该段时间是否被使用
        course: 一个list，用来维护该段时间所安排的课程
    """

    def __init__(self):
        """ Init Period """
        super(Period, self).__init__()
        self.isUsed = False
        self.course = []




class Schedule(object):
    """ 课程表类

    Attributes:
        day: 类所维护的课表，包含6个list，其中day[1] ~ day[5]表示周一到周五的课表
             每个day[i]中包含16个Period，每个Period表示一节课，
             其中day[1][1]表示周一第一节课
        season: 课程表所属的学季，0表示春／秋季，1表示夏／冬季
        classNumOfDay: 包含6个数，其中classNumOfDay[1]表示周一已安排的课程数量
        maxClassNumPerDay: 表示一天安排课程数量的上限，一般除非课程时间优先度为3，
                           否则每天课程数量不能大于该数
    """

    def __init__(self, season):
        super(Schedule,self).__init__()
        self.day = []  #schedul[1]~schedul[5]分别表示周一到周五，schedul[i][j],表示星期i的第j节课
        for i in range(6):
            self.day.append([])
            for j in range(16):
                p = Period()
                self.day[i].append(p)
        self.season = season
        self.classNumOfDay= []
        self.maxClassNumPerDay = 2
        for i in range(6):
            self.classNumOfDay.append(0);

    def add(self, courseList): #把课程加到课表里
        """ 根据每个课程可安排的时间，把课程列表里的课程加到课表里，从而生成课表 """
        turn = 4
        haveAddClass = False
        canConfilc = False
        while turn > 0 and len(courseList.list):
            successful = False
            for course in courseList.list:
                for i in range(1, 6):
                    for j in range(1, 16):
                        # print course.name, i, j, course.timeList[i][j], turn
                        if course.timeList[i][j] >= turn:
                            #print course.name, turn
                            if (turn >= 3) or ((not self.day[i][j].isUsed or canConfilc) and
                            (self.classNumOfDay[i] < self.maxClassNumPerDay) and
                            (j + course.length - 1 < 6 or
                            (j >= 6 and j + course.length - 1 < 11) or j > 10)):
                                isConfilic = False
                                for k in range(course.length):
                                    if self.day[i][j+k].isUsed and not canConfilc:
                                        isConfilic = True
                                        break
                                    for classAtThisTime in self.day[i][j+k].course:
                                        if classAtThisTime.isConflict(course):
                                            print course.name,"is conifilc to", classAtThisTime.name
                                            isConfilic = True
                                            break
                                    if isConfilic:
                                        break
                                if isConfilic:
                                    continue
                                for k in range(course.length):
                                    self.day[i][j+k].isUsed = True
                                    self.day[i][j+k].course.append(course)
                                    if course.season == 2:
                                        course.timeList[i][j+k] = 4
                                successful = True
                                self.classNumOfDay[i] += 1
                                courseList.list.remove(course)
                                break
                    if successful:
                        break
            if not successful:
                if  turn == 1 and not haveAddClass:
                    self.maxClassNumPerDay += 1
                    turn = 3
                    canConfilc = True
                    haveAddClass = True
                else:
                    turn -= 1

            #debug start

            if not turn:
                for course in courseList.list:
                    print "没有排完课程：", course.name

            #debug end

    def output_schedule(self, filename):
        """ 用排课结果生成execl课表 """
        file = xlwt.Workbook()
        table = file.add_sheet(u'sheet0',cell_overwrite_ok=True)
        if self.season == 0:
            table.write(0, 0, u'spring\u0020schedule')
        else:
            table.write(0, 0, u'summer\u0020schedule')
        for i in range(1, 6):
            table.write(1, 1+i*10, weekday_name[i-1])
            for j in range(1, 16):
                table.write(2 + (j-1)*3, 0, u'第%d节课'%j)
                k = 0
                for course in self.day[i][j].course:
                    table.write(2 + (j-1)*3, 1+ i*10 +k, course.name)
                    k += 1
        file.save(filename)







class CourseList(object):
    """docstring for CouserList"""
    def __init__(self):
        super(CourseList, self).__init__()
        self.list  = []

    def getCourseByXls(self, file = 'file.xls'):
        data = open_excel(file)
        table = data.sheets()[0]
        if SPRING_ROW >= 0:
            unicode_dic['Spring'] = table.cell(SPRING_ROW, SEASON_COL).value
        else:
            unicode_dic['Spring'] = None
        if SUMMER_ROW >= 0:
            unicode_dic['Summer'] = table.cell(SUMMER_ROW, SEASON_COL).value
        else:
            unicode_dic['Summer'] = None
        if SPRING_SUMMER_ROW >= 0:
            unicode_dic['Spring_Summer'] = table.cell(SPRING_SUMMER_ROW, SEASON_COL).value
        else:
            unicode_dic['Spring_Summer'] = None
        if REQUIRE_ROW >= 0:
            unicode_dic['Require'] = table.cell(REQUIRE_ROW, OPTIONAL_COL).value
        else:
            unicode_dic['Require'] = None

        #debug start
        #print unicode_dic['Spring'], unicode_dic['Summer']
        #debug end

        for i in range(klineStart, klineEnd + 1):
            #debug start
            #print i
            #debug end
            number = table.row(i)[NUMBER_COL].value
            name = table.row(i)[NAME_COL].value
            length = int(table.row(i)[LENGTH_COL].value)
            #debug start
            #print table.row(i)[SEASON_COL].value,unicode_dic['Summer']
            #debug end
            if table.row(i)[SEASON_COL].value == unicode_dic['Spring_Summer']:
                season = 2
            elif table.row(i)[SEASON_COL].value == unicode_dic['Spring']:
                season = 0
            elif table.row(i)[SEASON_COL].value == unicode_dic['Summer']:
                #print 1
                season = 1
            else:
                season = 3
            year =  int(table.row(i)[YEAR_COL].value)
            if table.row(i)[OPTIONAL_COL].value == unicode_dic['Require']:
                isOptional = False
            else:
                isOptional = True
            cal = []

            #预处理每个课程可排课的时间
            #基本要求为周一上午第一节不排，每天晚上不排，周五下午不排
            for m in range(6):
                cal.append([])
                for n in range(16):
                    cal[m].append(1)
            for m in range(6):
                cal[m][0] = 0
                #cal[m][1] = 0   #早上第一节不排课
            for m in range(6):
                for n in range(11, 16):
                    cal[m][n] = 0

            for k in range(int(table.row(i)[PRIORITY_COL].value)):
                temp_turn = int(table.row(i)[PRIORITY_COL+1+k*4].value)
                temp_day = int(table.row(i)[PRIORITY_COL+1+k*4+1].value)
                # print i, k
                temp_start = int(table.row(i)[PRIORITY_COL+1+k*4+2].value)
                temp_end = int(table.row(i)[PRIORITY_COL+1+k*4+3].value)
                if temp_day < 6:
                    for j in range(temp_start,temp_end+1):
                        cal[temp_day][j] = temp_turn
                else:
                    for weekday in cal:
                        for j in range(temp_start, temp_end+1):
                            weekday[j] = temp_turn

            for m in range(6, 11):
                cal[5][m] = 0

            course = Course(number, name, length, season, year, isOptional, cal)
            self.list.append(course)

    def getCourseByCourseList(self, anotherCourses, season):
        successful = True

        while successful:
            successful = False
            for course in anotherCourses.list:
                #debug start
                #print course.name,course.season, season
                #debug end
                if course.season == season:
                    self.list.append(course)
                    anotherCourses.list.remove(course)
                    successful = True
                elif course.season == 2 and course not in self.list:
                    self.list.append(course)
                    successful = True



class Course(object):
    """课程类"""
    def __init__(self, number, name, length, season, year, isOptional, timeList):
        super(Course,self).__init__()
        self.number = number    #课程编号
        self.name = name    #课程名称
        self.length = length    #课程时间
        self.season = season    #课程所在学季，0：春/秋，1：夏/冬，2：春夏/秋冬
        self.year = year;
        self.isOptional = isOptional    #课程是否是选修
        self.timeList = timeList    #课程可用时间表

    def isConflict(self, anotherCourse):
        if ((not self.isOptional or not anotherCourse.isOptional) and
        ((self.year > 0 and self.year == anotherCourse.year) or
        self.year == 1 or anotherCourse.year == 1)):
            return True
        if self.year > 0 and self.year == anotherCourse.year:
            return True

        return False



spring_schedule = Schedule(0)
summer_schedule = Schedule(1)

spring_course_list = CourseList()
spring_course_list.getCourseByXls("7.xls")

summer_course_list = CourseList()
summer_course_list.getCourseByCourseList(spring_course_list, 1)

#debug start
# for course in spring_course_list.list:
#     print course.name

#debug end



spring_schedule.add(spring_course_list)
summer_schedule.add(summer_course_list)

spring_schedule.output_schedule("7.spring.xls")
summer_schedule.output_schedule("7.summer.xls")
