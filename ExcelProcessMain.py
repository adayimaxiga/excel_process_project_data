#! /usr/bin/env python
# -*- coding:utf-8 –*-

import xlrd
import xlwt

# 打开
Workdata = xlrd.open_workbook('出版科学数据.xlsx')

Workdata2 = xlrd.open_workbook('出版科学被引.xlsx')
SheetName2 = Workdata2.sheet_names()
print(SheetName2)  # 打印名称
table2 = Workdata2.sheets()[0]
nrows2 = table2.nrows  # 行数
ncols2 = table2.ncols  # 列数

SheetName = Workdata.sheet_names()
print(SheetName)  # 打印名称
table = Workdata.sheets()[0]
nrows = table.nrows  # 行数
ncols = table.ncols  # 列数

# 写入
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
SheetElectric = book.add_sheet('表白我家傻敷敷', cell_overwrite_ok=True)
SheetElectric.write(0, 0, '作者')  # A1
SheetElectric.write(0, 1, '作者单位')  # AD
SheetElectric.write(0, 2, '标题')  # T1
SheetElectric.write(0, 3, '期刊')  # T1
SheetElectric.write(0, 4, '时间')  # T1
SheetElectric.write(0, 5, '文章数')  # T1
SheetElectric.write(0, 6, '按照姓名单位匹配相同作者')  # T1
ElecCount = 0

SheetMechnical = book.add_sheet('学猫叫', cell_overwrite_ok=True)
SheetMechnical.write(0, 0, '作者')  # A1
SheetMechnical.write(0, 1, '作者单位')  # AD
SheetMechnical.write(0, 2, '标题')  # T1
SheetMechnical.write(0, 3, '期刊')  # T1
SheetMechnical.write(0, 4, '时间')  # T1
SheetMechnical.write(0, 5, '文章数')  # T1
SheetMechnical.write(0, 6, '按照姓名匹配相同作者')  # T1

Sheet3 = book.add_sheet('喵喵喵喵喵', cell_overwrite_ok=True)
Sheet3.write(0, 0, '作者')  # A1
Sheet3.write(0, 1, '作者单位')  # AD
Sheet3.write(0, 2, '文章数')  # T1
Sheet3.write(0, 3, '作者及频次')  # T1

Sheet4 = book.add_sheet('傻敷敷不是智慧可爱', cell_overwrite_ok=True)
Sheet4.write(0, 0, '作者')  # A1
Sheet4.write(0, 1, '期刊')  # AD
Sheet4.write(0, 2, '被引频次')  # T1

Sheet5 = book.add_sheet('傻敷敷就是傻敷敷', cell_overwrite_ok=True)
Sheet5.write(0, 0, '作者')  # A1
Sheet5.write(0, 1, '作者单位')  # AD
Sheet5.write(0, 2, '文章数')  # T1
Sheet5.write(0, 3, '被引频次')  # T1


MechCount = 0


def ProcessExcelRow(Rawlist):
    Processedlist = ['header']
    Flag = False
    for i in range(len(Rawlist)):
        if (Rawlist[i] != ''):
            Processedlist.append(Rawlist[i])
        #            Processedlist.remove('header')
    return Processedlist


# 声明一个类
class MessegeOfRobomasterPerson:
    def __init__(self, Name, Office, title, Book, Date):
        self.name = Name
        self.office = Office
        self.title = title
        self.book = Book
        self.date = Date
        self.count = 0
        self.frequence =0
    def countSet(self, Count):
        self.count = Count
    def frequenceSet(self, frequence):
        self.frequence = frequence
    def TimeSet(self, Time):
        self.time = round(Time, 1)

    def print(self):
        print("作者：" + str(self.name))
        print("作者单位：" + str(self.office))
        print("标题：" + str(self.title))
        print("期刊：" + str(self.title))

    def ReturnDate(self):
        return self.date
    def ReturnFrequence(self):
        return self.frequence
    def ReturnBook(self):
        return self.book

    def ReturnName(self):
        return self.name

    def ReturnOffice(self):
        return self.office

    def Returntitle(self):
        return self.title

    def Returncount(self):
        return self.count

    def __cmp__(self, other):
        if self.name != other.name:
            return self.name < other.name
        elif self.name == other.name:
            return self.office < other.name

    def __lt__(self, other):
        if self.name != other.name:
            return self.name < other.name
        elif self.name == other.name:
            if self.office != other.office:
                return self.office < other.office
            elif self.office == other.office:
                return self.title < other.title

    def __eq__(self, other):
        if (self.name == other.name) and (self.office == other.office) and (self.title == other.title):
            return True
        else:
            return False


def safe_int(num):
    try:
        return int(num)
    except ValueError:
        result = []
        for c in num:
            if not ('0' <= c <= '9'):
                break
            result.append(c)
        if len(result) == 0:
            return 0
        return int(''.join(result))


def ConvertToHours(StringTime):
    BufferTimeHour = StringTime[0:2]
    BufferTimeMinute = StringTime[3:5]
    BufferTimeHourInt = safe_int(BufferTimeHour)
    BufferTimeMinuteInt = safe_int(BufferTimeMinute)
    return BufferTimeHourInt + BufferTimeMinuteInt / 60


def isNum(value):
    try:
        value + 1
    except TypeError:
        return False
    else:
        return True


person = []
person2 = []
if __name__ == "__main__":
    # 全局保存
    Count = 0
    FlagStartRecord = False
    namename = ""
    officeoffice = ""
    titletitle = ""
    bookbook = ""
    datedate = ""
    for i in range(nrows):
        # 循环初始变量
        FlagNameUpdate = False

        # 获取原始行列
        ListTempDeleteTemp = ProcessExcelRow(table.row_values(i))
        if (len(ListTempDeleteTemp) > 2):
            if (ListTempDeleteTemp[1] == "A1"):
                if (Count > 0):
                    person.append(MessegeOfRobomasterPerson(namename, officeoffice, titletitle, bookbook, datedate))
                    namename = ""
                    officeoffice = ""
                    titletitle = ""
                    bookbook = ""
                    datedate = ""
                    FlagNameUpdate = True
                namename = ListTempDeleteTemp[2]
                Count = Count + 1
            if (ListTempDeleteTemp[1] == "AD"):
                officeoffice = ListTempDeleteTemp[2]
            if (ListTempDeleteTemp[1] == "T1"):
                titletitle = ListTempDeleteTemp[2]
            if (ListTempDeleteTemp[1] == "JF"):
                bookbook = ListTempDeleteTemp[2]
            if (ListTempDeleteTemp[1] == "YR"):
                datedate = ListTempDeleteTemp[2]
    person.append(MessegeOfRobomasterPerson(namename, officeoffice, titletitle, bookbook, datedate))
    print(Count)
    person.sort()
    flagNew = True
    temp = person[0]
    i = 1
    while i < len(person):  # 删除重复部分
        if (person[i] == temp):
            person.pop(i)
        else:
            temp = person[i]
            i += 1
    nametemp = person[0].name
    officetemp = person[0].office
    count = 1
    number = 0
    for i in range(1, len(person)):
        if (person[i].name == nametemp) and (officetemp == person[i].office):
            count += 1
        else:
            nametemp = person[i].name
            officetemp = person[i].office
            number = count
            while (count >= 1):
                person[i - count].countSet(number)
                count -= 1
            count = 1
    person[i].countSet(count)

    for i in range(len(person)):
        SheetElectric.write(i + 1, 0, person[i].ReturnName())
        SheetElectric.write(i + 1, 1, person[i].ReturnOffice())
        SheetElectric.write(i + 1, 2, person[i].Returntitle())
        SheetElectric.write(i + 1, 3, person[i].ReturnBook())
        SheetElectric.write(i + 1, 4, person[i].ReturnDate())
        if (person[i].Returncount() != 0):
            SheetElectric.write(i + 1, 5, person[i].Returncount())
    # sheet2处理
    i=0
    while i < len(person):  # 删除无office数据
        if (person[i].ReturnOffice() == ""):
            person.pop(i)
        else:
            person[i].countSet(0)
            i+=1
    nametemp = person[0].name
    count =1
    number = 0
    for i in range(1, len(person)):
        if (person[i].name == nametemp):
            count += 1
        else:
            nametemp = person[i].name
            number = count
            while (count >= 1):
                person[i - count].countSet(number)
                count -= 1
            count = 1
    person[i].countSet(count)
    for i in range(len(person)):
        SheetMechnical.write(i + 1, 0, person[i].ReturnName())
        SheetMechnical.write(i + 1, 1, person[i].ReturnOffice())
        SheetMechnical.write(i + 1, 2, person[i].Returntitle())
        SheetMechnical.write(i + 1, 3, person[i].ReturnBook())
        SheetMechnical.write(i + 1, 4, person[i].ReturnDate())
        if (person[i].Returncount() != 0):
            SheetMechnical.write(i + 1, 5, person[i].Returncount())

    #sheet3
    i = 1
    tempnamename = person[0].ReturnName()
    while i < len(person):  # 删除重复部分
        if (person[i].ReturnName() == tempnamename):
            person.pop(i)
        else:
            tempnamename = person[i].ReturnName()
            i += 1

    for i in range(len(person)):
        Sheet3.write(i + 1, 0, person[i].ReturnName())
        Sheet3.write(i + 1, 1, person[i].ReturnOffice())
        Sheet3.write(i + 1, 2, person[i].Returncount())
    #读取excel2数据
    namename2 = ""
    bookbook2 = ""
    countcount2 =0

    for i in range(1,nrows2):
        # 获取原始行列
        ListTempDeleteTemp2 = ProcessExcelRow(table2.row_values(i))
        if (len(ListTempDeleteTemp2) > 2):
            namename2 = ListTempDeleteTemp2[2]
            bookbook2 = ListTempDeleteTemp2[3]
            countcount2 = ListTempDeleteTemp2[4]

        person2.append(MessegeOfRobomasterPerson(namename2,"","",bookbook2,""))
        person2[i-1].countSet(countcount2)
    #print(person2[0].ReturnName())
    #person2.sort()
    #print(person2[0].ReturnName())
    #sheet4
    temp22=person2[0].ReturnName()
    i=1
    while i < len(person2):  #删除重复部分
        if (person2[i].ReturnName() == temp22):
            person2[i-1].countSet(person2[i-1].Returncount()+person2[i].Returncount())
            person2.pop(i)
        else:
            temp22 = person2[i].ReturnName()
            i += 1
    for i in range(len(person2)):
        Sheet4.write(i + 1, 0, person2[i].ReturnName())
        Sheet4.write(i + 1, 1, person2[i].ReturnBook())
        Sheet4.write(i + 1, 2, person2[i].Returncount())

    for i in range(len(person)):
        for j in range(len(person2)):
            if(person[i].ReturnName()==person2[j].ReturnName()):
                person[i].frequenceSet(person2[j].Returncount())
    j=0
    for i in range(len(person)):
        if(person[i].ReturnFrequence()!=0):
            Sheet5.write(j + 1, 0, person[i].ReturnName())
            Sheet5.write(j + 1, 1, person[i].ReturnOffice())
            Sheet5.write(j + 1, 2, person[i].Returncount())
            Sheet5.write(j + 1, 3, person[i].ReturnFrequence())
            Sheet5.write(j + 1, 4, person[i].ReturnBook())
            j+=1

book.save(r'傻敷敷看过来.xls')
