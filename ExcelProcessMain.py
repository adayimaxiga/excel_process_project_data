#! /usr/bin/env python
# -*- coding:utf-8 –*-

import xlrd
import xlwt
import ngender
#import re

import sys

file_name_output=input("请从输入期刊名，出版编辑类（中国出版，中国科技期刊研究，出版发行研究，出版科学，现代出版，科技与出版，编辑之友）新闻与传媒类（国际新闻界，当代传播，新闻与传播研究，新闻大学，新闻界，新闻记者，现代传播）")

all_chubanlei = {'中国出版','中国科技期刊研究','出版发行研究','出版科学','现代出版','科技与出版','编辑之友'}
chubanlei = False

for items in all_chubanlei:
    if(items==file_name_output):
        chubanlei = True

#file_name_output = '新闻界'
#测试代码
file_name_daochu = file_name_output+'导出2000-2017'
file_name_beiyin = file_name_output+'被引2000-2017'

# 打开
Workdata_daochu = xlrd.open_workbook(file_name_daochu+'.xlsx')
Workdata_beiyin = xlrd.open_workbook(file_name_beiyin+'.xlsx')

sheetname_daochu = Workdata_daochu.sheet_names()
sheetname_beiyin = Workdata_beiyin.sheet_names()

#print(sheetname_beiyin)

#处理导出xlsx
table_daochu  = Workdata_daochu.sheets()[0]
nrows_daochu  = table_daochu .nrows  # 行数
ncols_daochu  = table_daochu .ncols  # 列数
#处理被引xlsx
table_beiyin  = Workdata_beiyin.sheets()[1]
nrows_beiyin  = table_beiyin .nrows  # 行数
ncols_beiyin  = table_beiyin .ncols  # 列数


#导出文件
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
Sheet_daochu_raw = book.add_sheet(file_name_output+'数据导出（筛选提取后）', cell_overwrite_ok=True)
Sheet_daochu_raw.write(0, 0, 'SrcDatabase')  # A1
Sheet_daochu_raw.write(0, 1, 'Title')  # A1
Sheet_daochu_raw.write(0, 2, 'Author')  # A1
Sheet_daochu_raw.write(0, 3, 'Organ')  # A1
Sheet_daochu_raw.write(0, 4, 'Source')  # A1
Sheet_daochu_raw.write(0, 5, 'Keyword')  # A1
Sheet_daochu_raw.write(0, 6, 'Summary')  # A1
Sheet_daochu_raw.write(0, 7, 'PubTime')  # A1
Sheet_daochu_raw.write(0, 8, 'FirstDuty')  # A1
Sheet_daochu_raw.write(0, 9, 'Year')  # A1
Sheet_daochu_raw.write(0, 10, 'Volume')  # A1
Sheet_daochu_raw.write(0, 11, 'Period')  # A1
Sheet_daochu_raw.write(0, 12, 'PageCount')  # A1
Sheet_daochu_raw.write(0, 13, 'CLC')  # A1

Sheet_beiyin_raw = book.add_sheet(file_name_output+'数据被引（原始数据）', cell_overwrite_ok=True)
Sheet_beiyin_raw.write(0, 0, 'Number')  # A1
Sheet_beiyin_raw.write(0, 1, 'Title')  # A1
Sheet_beiyin_raw.write(0, 2, 'Author')  # A1
Sheet_beiyin_raw.write(0, 3, 'Source')  # A1
Sheet_beiyin_raw.write(0, 4, 'PubTime')  # A1
Sheet_beiyin_raw.write(0, 5, 'DataBase')  # A1
Sheet_beiyin_raw.write(0, 6, 'Reference')  # A1
Sheet_beiyin_raw.write(0, 7, 'Download')  # A1

Sheet_pipei = book.add_sheet(file_name_output+'（匹配后）', cell_overwrite_ok=True)
Sheet_pipei.write(0, 0, 'Number')  # A1
Sheet_pipei.write(0, 1, 'Title')  # A1
Sheet_pipei.write(0, 2, 'Author')  # A1
Sheet_pipei.write(0, 3, 'Source')  # A1
Sheet_pipei.write(0, 4, 'PubTime')  # A1
Sheet_pipei.write(0, 5, 'DataBase')  # A1
Sheet_pipei.write(0, 6, 'Reference')  # A1
Sheet_pipei.write(0, 7, 'Download')  # A1
Sheet_pipei.write(0, 8, '题名')  # A1
Sheet_pipei.write(0, 9, '作者')  # A1
Sheet_pipei.write(0, 10, '单位')  # A1
Sheet_pipei.write(0, 11, '关键词')  # A1
Sheet_pipei.write(0, 12, '摘要')  # A1
Sheet_pipei.write(0, 13, '基金')  # A1
Sheet_pipei.write(0, 14, '页码')  # A1
Sheet_pipei.write(0, 15, '页数')  # A1
Sheet_pipei.write(0, 16, '权值_作者')  # A1
Sheet_pipei.write(0, 17, '权值_单位')  # A1
Sheet_pipei.write(0, 18, '权值_基金')  # A1
Sheet_pipei.write(0, 19, '第一责任人')  # A1
Sheet_pipei.write(0, 20, '性别估计')  # A1
Sheet_pipei.write(0, 21, '性别估计概率')  # A1

#存储所有导出的类
data_daochu=[]
#存储宝宝要的类
data_daochu_deleteuseless=[]
#存储所有被引的类
data_beiyin=[]

#声明一个存储数据的类
class items_daochu:
    SrcDatabase = ''
    Title = ''
    Author = ''
    Organ = ''
    Source = ''
    Keyword = ''
    Summary = ''
    PubTime = ''
    FirstDuty = ''
    Fund = ''
    Year = ''
    Volume = ''
    Period = ''
    PageCount = ''
    CLC = ''
    PageCal = 0
    def __init__(self, SrcDatabase, Title, Author, Organ, Source,Keyword,Summary,PubTime,FirstDuty,Fund,Year,Volume,Period,PageCount,CLC,PageCal):
        self.SrcDatabase = SrcDatabase
        self.Title = Title
        self.Author = Author
        self.Organ = Organ
        self.Source = Source
        self.Keyword = Keyword
        self.Summary = Summary
        self.PubTime = PubTime
        self.FirstDuty = FirstDuty
        self.Fund = Fund
        self.Year = Year
        self.Volume = Volume
        self.Period = Period
        self.PageCount = PageCount
        self.CLC = CLC
        self.PageCal = PageCal


    def SrcDatabase(self):
        return self.SrcDatabase
    def Title(self):
        return self.Title
    def Author(self):
        return self.Author
    def Organ(self):
        return self.Organ
    def Source(self):
        return self.Source
    def Keyword(self):
        return self.Keyword
    def Summary(self):
        return self.Summary
    def PubTime(self):
        return self.PubTime
    def FirstDuty(self):
        return self.FirstDuty
    def Fund(self):
        return self.Fund
    def Year(self):
        return self.Year
    def Volume(self):
        return self.Volume
    def Period(self):
        return self.Period
    def PageCount(self):
        return self.PageCount
    def CLC(self):
        return self.CLC

    def print(self):
        print("SrcDatabase"+self.SrcDatabase)
        print("Title" + self.Title)
        print("Author" + self.Author)
        print("Organ" + self.Organ)
        print("Source" + self.Source)
        print("Keyword" + self.Keyword)
        print("Summary" + self.Summary)
        print("PubTime" + self.PubTime)
        print("FirstDuty" + self.FirstDuty)
        print("Fund" + self.Fund)
        print("Year" + self.Year)
        print("Volume" + self.Volume)
        print("Period" + self.Period)
        print("PageCount" + self.PageCount)
        print("CLC" + self.CLC)

#声明一个存储数据的类
class items_beiyin:
    Number = ''
    Title = ''
    Author = ''
    Source = ''
    PubTime = ''
    Database = ''
    Reference = ''
    Download = ''

    def __init__(self, Number, Title, Author, Source, PubTime,Database,Reference,Download):
        self.Number = Number
        self.Title = Title
        self.Author = Author
        self.Source = Source
        self.PubTime = PubTime
        self.Database = Database
        self.Reference = Reference
        self.Download = Download

    def print(self):
        print("Number"+str(self.Number))
        print("Title" + str(self.Title))
        print("Author" + str(self.Author))
        print("Source" + str(self.Source))
        print("PubTime" + str(self.PubTime))
        print("Database" + str(self.Database))
        print("Reference" + str(self.Reference))
        print("Download" + str(self.Download))

def ProcessExcelRow(Rawlist):
    Processedlist = []
    Flag = False
    for i in range(len(Rawlist)):
        if (Rawlist[i] != ''):
            Processedlist.append(Rawlist[i])
        #            Processedlist.remove('header')
    return Processedlist

def ProcessExcelRow_notdeleteblank(Rawlist):
    Processedlist = []
    Flag = False
    for i in range(len(Rawlist)):
        #if (Rawlist[i] != ''):
            Processedlist.append(Rawlist[i])
        #            Processedlist.remove('header')
    return Processedlist

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
def itemsprocess(itemsnow):

    flag_while= True
    i=0
    titlenow=''
    SrcDatabase = ''
    Title = ''
    Author = ''
    Organ = ''
    Source = ''
    Keyword = ''
    Summary = ''
    PubTime = ''
    FirstDuty = ''
    Fund = ''
    Year = ''
    Volume = ''
    Period = ''
    PageCount = ''
    CLC = ''
    PageCal = 0
    if(len(itemsnow)>2):
        # 处理换行但是没有新条款的情况。
        while flag_while:
            split = itemsnow[i].find('-')
            if(split==-1):
                itemsnow[i-1]=itemsnow[i-1]+itemsnow[i]
                del itemsnow[i]
                #print('haha')
            i=i+1
            if(i>=len(itemsnow)):
                flag_while = False

        for i in range(len(itemsnow)):
            split = itemsnow[i].find('-')
            titlenow = itemsnow[i][0:split]
            if(titlenow == 'SrcDatabase'):
                maohao = itemsnow[i].find(':',split)
                SrcDatabase = itemsnow[i][maohao+1:].lstrip(' ')
                #print(SrcDatabase)
            elif(titlenow == 'Title'):
                maohao = itemsnow[i].find(':', split)
                Title = itemsnow[i][maohao+1:].lstrip(' ')
            elif (titlenow == 'Author'):
                maohao = itemsnow[i].find(':', split)
                Author = itemsnow[i][maohao + 1:].lstrip(' ')
            elif (titlenow == 'Organ'):
                maohao = itemsnow[i].find(':', split)
                Organ = itemsnow[i][maohao + 1:].lstrip(' ')
            elif (titlenow == 'Source'):
                maohao = itemsnow[i].find(':', split)
                Source = itemsnow[i][maohao + 1:].lstrip(' ')
            elif (titlenow == 'Keyword'):
                maohao = itemsnow[i].find(':', split)
                Keyword = itemsnow[i][maohao + 1:].lstrip(' ')
            elif (titlenow == 'Summary'):
                maohao = itemsnow[i].find(':', split)
                Summary = itemsnow[i][maohao + 1:].lstrip(' ')
            elif (titlenow == 'PubTime'):
                maohao = itemsnow[i].find(':', split)
                PubTime = itemsnow[i][maohao + 1:].lstrip(' ')
            elif (titlenow == 'FirstDuty'):
                maohao = itemsnow[i].find(':', split)
                FirstDuty = itemsnow[i][maohao + 1:].lstrip(' ').rstrip(';')
            elif (titlenow == 'Fund'):
                maohao = itemsnow[i].find(':', split)
                Fund = itemsnow[i][maohao + 1:].lstrip(' ')
            elif (titlenow == 'Year'):
                maohao = itemsnow[i].find(':', split)
                Year = itemsnow[i][maohao + 1:].lstrip(' ')
            elif (titlenow == 'Volume'):
                maohao = itemsnow[i].find(':', split)
                Volume = itemsnow[i][maohao + 1:].lstrip(' ')
            elif (titlenow == 'Period'):
                maohao = itemsnow[i].find(':', split)
                Period = itemsnow[i][maohao + 1:].lstrip(' ')
            elif (titlenow == 'PageCount'):
                maohao = itemsnow[i].find(':', split)
                PageCount = itemsnow[i][maohao + 1:].lstrip(' ').rstrip(' ')
            elif (titlenow == 'CLC'):
                maohao = itemsnow[i].find(':', split)
                CLC = itemsnow[i][maohao + 1:].lstrip(' ')
        temp_loc=PageCount.find('-')
        temp_loc_plus = PageCount.find('+')
        first_number =0
        second_number =0
        if(temp_loc!=-1):
            first_number=safe_int(PageCount[0:temp_loc])
            if(temp_loc_plus==-1):
                second_number = safe_int(PageCount[temp_loc+1:])
            else:
                second_number = safe_int(PageCount[temp_loc+1:temp_loc_plus])+1
            PageCal =second_number -first_number   + 1
        else:
            PageCal = 1

    return items_daochu(SrcDatabase, Title, Author, Organ, Source,Keyword,Summary,PubTime,FirstDuty,Fund,Year,Volume,Period,PageCount,CLC,PageCal)

def check_contain_chinese(check_str):
        for c in check_str:
            if not ('\u4e00' <= c <= '\u9fa5'):
                return False
        return True

if __name__ == "__main__":
    grid = [['.', '.', '.', '.', '.', '.'],
            ['.', 'O', 'O', '.', '.', '.'],
            ['O', 'O', 'O', 'O', '.', '.'],
            ['O', 'O', 'O', 'O', 'O', '.'],
            ['.', 'O', 'O', 'O', 'O', 'O'],
            ['O', 'O', 'O', 'O', 'O', '.'],
            ['O', 'O', 'O', 'O', '.', '.'],
            ['.', 'O', 'O', '.', '.', '.'],
            ['.', '.', '.', '.', '.', '.']]
    for i in range(6):
        for x in range(10):
            if x == 9:
                print('\r')
            else:
                print(grid[x][i], end=' ')
    flag_new = False   #新数据flag

    items_now=[]
    #读取每一行
    for i in range(nrows_daochu):
        ListTempDeleteTemp = ProcessExcelRow(table_daochu.row_values(i))
        #删除空白
        if (len(ListTempDeleteTemp) > 0):

            if(ListTempDeleteTemp[0][0:11]=='SrcDatabase'):
                flag_new = True
            #数据更新完毕，开始处理
            if(flag_new):
                #print(items_now)
                data_daochu.append(itemsprocess(items_now))
                items_now = []
                flag_new = False
                pass
            items_now.append(ListTempDeleteTemp[0])
    #筛选条件
    for i in range(len(data_daochu)):
        if(data_daochu[i].Author !=''):
            if (data_daochu[i].Keyword != ''):
                if(file_name_output == '现代出版'):
                    if((data_daochu[i].Source.find(file_name_output)!=-1) or(data_daochu[i].Source.find('大学出版')!=-1)):
                        data_daochu_deleteuseless.append(data_daochu[i])
                else:
                    if (data_daochu[i].Source.find(file_name_output) != -1):
                        data_daochu_deleteuseless.append(data_daochu[i])

    print("导出数据筛选结果 before : " ,len(data_daochu) , "after : ", len(data_daochu_deleteuseless))

    for i in range(len(data_daochu_deleteuseless)):
        Sheet_daochu_raw.write(i + 1, 0, data_daochu_deleteuseless[i].SrcDatabase)
        Sheet_daochu_raw.write(i + 1, 1, data_daochu_deleteuseless[i].Title)
        Sheet_daochu_raw.write(i + 1, 2, data_daochu_deleteuseless[i].Author)
        Sheet_daochu_raw.write(i + 1, 3, data_daochu_deleteuseless[i].Organ)
        Sheet_daochu_raw.write(i + 1, 4, data_daochu_deleteuseless[i].Source)
        Sheet_daochu_raw.write(i + 1, 5, data_daochu_deleteuseless[i].Keyword)
        Sheet_daochu_raw.write(i + 1, 6, data_daochu_deleteuseless[i].Summary)
        Sheet_daochu_raw.write(i + 1, 7, data_daochu_deleteuseless[i].PubTime)
        Sheet_daochu_raw.write(i + 1, 8, data_daochu_deleteuseless[i].FirstDuty)
        Sheet_daochu_raw.write(i + 1, 9, data_daochu_deleteuseless[i].Year)
        Sheet_daochu_raw.write(i + 1, 10, data_daochu_deleteuseless[i].Volume)
        Sheet_daochu_raw.write(i + 1, 11, data_daochu_deleteuseless[i].Period)
        Sheet_daochu_raw.write(i + 1, 12, data_daochu_deleteuseless[i].PageCount)
        Sheet_daochu_raw.write(i + 1, 13, data_daochu_deleteuseless[i].CLC)





        #开始处理被引数据
    for i in range(nrows_beiyin):
        ListTempDeleteTemp = ProcessExcelRow_notdeleteblank(table_beiyin.row_values(i))
        #print(ListTempDeleteTemp)
        Number = ''
        Title = ''
        Author = ''
        Source = ''
        PubTime = ''
        Database = ''
        Reference = ''
        Download = ''
        Number = ListTempDeleteTemp[0]
        Title = ListTempDeleteTemp[1]
        Author = ListTempDeleteTemp[2]
        Source = ListTempDeleteTemp[3]
        PubTime = ListTempDeleteTemp[4]
        Database = ListTempDeleteTemp[5]
        Reference = ListTempDeleteTemp[6]
        if(chubanlei==False):
            Download = ListTempDeleteTemp[7]
        data_beiyin.append(items_beiyin(Number, Title, Author, Source, PubTime,Database,Reference,Download))
        #data_beiyin[i].print()

    for i in range(len(data_beiyin)):
        Sheet_beiyin_raw.write(i + 1, 0, data_beiyin[i].Number)
        Sheet_beiyin_raw.write(i + 1, 1, data_beiyin[i].Title)
        Sheet_beiyin_raw.write(i + 1, 2, data_beiyin[i].Author)
        Sheet_beiyin_raw.write(i + 1, 3, data_beiyin[i].Source)
        Sheet_beiyin_raw.write(i + 1, 4, data_beiyin[i].PubTime)
        Sheet_beiyin_raw.write(i + 1, 5, data_beiyin[i].Database)
        Sheet_beiyin_raw.write(i + 1, 6, data_beiyin[i].Reference)
        Sheet_beiyin_raw.write(i + 1, 7, data_beiyin[i].Download)
    count =0
    for i in range(len(data_beiyin)):
        for j in range(len(data_daochu_deleteuseless)):
            if(data_beiyin[i].Title.find(data_daochu_deleteuseless[j].Title)!=-1):

                #if (data_beiyin[i].Author.find(data_daochu_deleteuseless[j].Author)!=-1):
                    Sheet_pipei.write(count + 1, 0, data_beiyin[i].Number)
                    Sheet_pipei.write(count + 1, 1, data_beiyin[i].Title)
                    Sheet_pipei.write(count + 1, 2, data_beiyin[i].Author)
                    Sheet_pipei.write(count + 1, 3, data_beiyin[i].Source)
                    Sheet_pipei.write(count + 1, 4, data_beiyin[i].PubTime)
                    Sheet_pipei.write(count + 1, 5, data_beiyin[i].Database)
                    Sheet_pipei.write(count + 1, 6, data_beiyin[i].Reference)
                    Sheet_pipei.write(count + 1, 7, data_beiyin[i].Download)
                    Sheet_pipei.write(count + 1, 8, data_daochu_deleteuseless[j].Title)  # A1
                    Sheet_pipei.write(count + 1, 9, data_daochu_deleteuseless[j].Author)  # A1
                    Sheet_pipei.write(count + 1, 10,data_daochu_deleteuseless[j].Organ)  # A1
                    Sheet_pipei.write(count + 1, 11,data_daochu_deleteuseless[j].Keyword)  # A1
                    Sheet_pipei.write(count + 1, 12,data_daochu_deleteuseless[j].Summary)  # A1
                    Sheet_pipei.write(count + 1, 13,data_daochu_deleteuseless[j].Fund)  # A1
                    Sheet_pipei.write(count + 1, 14,data_daochu_deleteuseless[j].PageCount)  # A1
                    Sheet_pipei.write(count + 1, 19,data_daochu_deleteuseless[j].FirstDuty)  # A1


                    if(check_contain_chinese(data_daochu_deleteuseless[j].FirstDuty)):
                        gender_this=ngender.guess(data_daochu_deleteuseless[j].FirstDuty)


                        if(gender_this[0] == 'male'):
                            Sheet_pipei.write(count + 1, 20, '男')
                            Sheet_pipei.write(count + 1, 21,gender_this[1])
                        elif(gender_this[0] == 'female'):
                            Sheet_pipei.write(count + 1, 20, '女')
                            Sheet_pipei.write(count + 1, 21, gender_this[1])

                    Author_temp = data_beiyin[i].Author.rstrip(';')
                    count_fen=0
                    if(Author_temp.find('课题组')!=-1):
                        count_fen=10
                    else:
                        for ch in Author_temp:
                            if(ch == ';'):
                                count_fen = count_fen +1
                            elif(ch == ','):
                                count_fen = count_fen + 1
                    if(count_fen>0):
                        Sheet_pipei.write(count + 1, 16, '1')  # A1
                    else:
                        Sheet_pipei.write(count + 1, 16, '0')  # A1

                    if(data_daochu_deleteuseless[j].Fund!=''):
                        Sheet_pipei.write(count + 1, 18, '1')  # A1
                    else:
                        Sheet_pipei.write(count + 1, 18, '0')  # A1
                    Sheet_pipei.write(count + 1, 15, data_daochu_deleteuseless[j].PageCal)  # A1
                    count = count + 1



book.save(file_name_output+r'傻敷敷看过来.xls')









'''
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

    grid = [['.', '.', '.', '.', '.', '.'],
            ['.', 'O', 'O', '.', '.', '.'],
            ['O', 'O', 'O', 'O', '.', '.'],
            ['O', 'O', 'O', 'O', 'O', '.'],
            ['.', 'O', 'O', 'O', 'O', 'O'],
            ['O', 'O', 'O', 'O', 'O', '.'],
            ['O', 'O', 'O', 'O', '.', '.'],
            ['.', 'O', 'O', '.', '.', '.'],
            ['.', '.', '.', '.', '.', '.']]
    for i in range(6):
        for x in range(10):
            if x == 9:
                print('\r')
            else:
                print(grid[x][i],end=' ')

    tempnum = int(input("傻敷敷告诉我一个数字"))
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
        if(len(ListTempDeleteTemp2)==5):
            countcount2 =0
        if(countcount2<tempnum):
            countcount2 =0
        person2.append(MessegeOfRobomasterPerson(namename2,"","",bookbook2,""))
        person2[i-1].countSet(countcount2)
    #print(person2[0].ReturnName())
    person2.sort()
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
'''