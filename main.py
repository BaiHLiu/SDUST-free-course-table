'''
Description: 导出空闲时间
Author: Catop
Date: 2021-09-28 23:45:08
LastEditTime: 2021-09-29 12:02:13
'''
import requests
import xlrd
import xlwt
import json

########################################
#查询时间
QUERY_DATES = [
    ['2021-09-20','2021-09-26'],
    ['2021-09-27','2021-10-03'],
    ['2021-10-04','2021-10-10'],
    ['2021-10-11','2021-10-17'],
    ['2021-10-18','2021-10-24'],
    ['2021-10-25','2021-10-31'],
    ['2021-11-01','2021-11-07'],
    ['2021-10-08','2021-10-14'],
    ['2021-10-15','2021-10-21'],
    ['2021-10-22','2021-10-28'],
    ['2021-10-29','2021-12-05'],
    ['2021-12-06','2021-12-12'],
    ['2021-12-13','2021-12-19'],
    ['2021-12-20','2021-12-26'],
    ['2021-12-27','2022-01-02'],
    ['2022-01-01','2022-01-09'],
    ['2022-01-10','2022-01-16'],
]

#API地址
API_HOST_1 = 'http://192.168.111.167:8080'
API_HOST_2 = 'http://192.168.111.167:8081'
#请求Headers
headers = {
        'Accept-Encoding':'gzip, deflate, br',
        'deviceSn':'aa3d119f9eda',
        'X-Consumer-Custom-ID':'sdust'
    }
########################################


def getFreeCourse(stuList,startDate,endDate):
    """获取指定时间段每天每节课无课人员"""
    # 初始化列表
    busyCourseList = []
    freeCourseList = []
    for day in range(0,7):
        dayCourse = []
        for slot in range(0,10):
            slotCourse = []
            dayCourse.append(slotCourse)
        busyCourseList.append(dayCourse)

    for day in range(0,7):
        dayCourse = []
        for slot in range(0,10):
            slotCourse = []
            dayCourse.append(slotCourse)
        freeCourseList.append(dayCourse)


    # 每次查一个学生的一周课表
    for stuInfo in stuList:
        print(stuInfo)
        stuId = stuInfo[0]
        #print(stuId)
        url = f"{API_HOST_2}/api/v3/students/{stuId}/timetable-items?start={startDate}&end={endDate}"

        try:
            req = requests.get(url, headers=headers)
            resJson = json.loads(req.text)
            
            for course in resJson:
                dayOfWeek = course['dayOfWeek']
                slotOfDay = course['slotOfDay']
                
                busyCourseList[dayOfWeek-1][slotOfDay-1].append(stuId)
        except:
            continue

        
    # 整理成空闲列表
    for day in range(0,7):
        for slot in range(0,10):
            for stuInfo in stuList:
                stuId = stuInfo[0]
                stuName = stuInfo[1]
                if not(stuId in busyCourseList[day][slot]):
                    freeCourseList[day][slot].append(stuName)

    return freeCourseList




def getStudentsInfo():
    """获取导入列表的学生信息"""
    dataIn = xlrd.open_workbook("./in.xls")
    table = dataIn.sheets()[0]
    nrows = table.nrows

    stuList = []

    for idx in range(0,nrows):
    
        stuId = table.cell_value(idx, 0)
        url = f"{API_HOST_2}/api/v3/users?uids={stuId}"
        req = requests.get(url, headers=headers)
        stuName = json.loads(req.text)[0]['fullName']

        stuList.append([
            stuId,
            stuName
        ])

        print(stuId+"-"+stuName)

    return stuList


def work():
    workbook = xlwt.Workbook(encoding = 'utf-8')
    style = xlwt.XFStyle()
    style.alignment.wrap = 1  #设置自动换行
    stuList = getStudentsInfo()
    # 各周时间
    for queryRange in QUERY_DATES:
        print(queryRange)
        worksheet = workbook.add_sheet(f"{queryRange[0]}_{queryRange[1]}")
        freeCourseList = getFreeCourse(stuList, queryRange[0], queryRange[1])
        for day in range(0,7):
            for slot in range(0,10):

                worksheet.col(day).width = 256 * 50  # Set the column width
                worksheet.write(slot,day, ','.join(freeCourseList[day][slot]), style)
                
                
                

                

    workbook.save('out.xls')

if __name__== "__main__":
    #print(getStudentsInfo())
    #print(getFreeCourse(getStudentsInfo(), QUERY_DATES[0][0], QUERY_DATES[0][1]))
    work()
    #(getFreeCourse(getStudentsInfo(), QUERY_DATES[1][0], QUERY_DATES[1][1]))