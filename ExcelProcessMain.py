#! /usr/bin/env python
# -*- coding:utf-8 –*-

import xlrd
import xlwt
import ngender

file_name_output=input("请从输入期刊名，出版编辑类（中国出版，中国科技期刊研究，出版发行研究，出版科学，现代出版，科技与出版，编辑之友，编辑学报）新闻与传媒类（国际新闻界，当代传播，新闻与传播研究，新闻大学，新闻界，新闻记者，现代传播）")
shuangyiliu ={'北京大学','中国人民大学','清华大学','北京航空航天大学','北京理工大学','中国农业大学','北京师范大学','中央民族大学','南开大学','天津大学','大连理工大学','吉林大学','哈尔滨工业大学','复旦大学','同济大学','上海交通大学','华东师范大学','南京大学','东南大学','浙江大学','中国科学技术大学','厦门大学','山东大学','中国海洋大学','武汉大学','华中科技大学','中南大学','中山大学','华南理工大学','四川大学','重庆大学','电子科技大学','西安交通大学','西北工业大学','兰州大学','国防科技大学','东北大学','郑州大学','湖南大学','云南大学','西北农林科技大学','新疆大学'}
all_chubanlei = {'中国出版','中国科技期刊研究','出版发行研究','出版科学','现代出版','科技与出版','编辑之友','编辑学报'}
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
Sheet_pipei.write(0, 22, '性别权值')  # A1
Sheet_pipei.write(0, 23, '发表月份权值')  # A1
Sheet_pipei.write(0, 24, '摘要题目权重')  # A1
Sheet_pipei.write(0, 25, '作者单位权重')  # A1
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
                    #print(data_daochu_deleteuseless[j].PubTime)
                    yuefen_start = data_daochu_deleteuseless[j].PubTime.find('-')
                    yuefen_n = safe_int(data_daochu_deleteuseless[j].PubTime[yuefen_start+1:yuefen_start+3])
                    if(yuefen_n<4):
                        Sheet_pipei.write(count + 1, 23, '0')
                    elif(yuefen_n<7):
                        Sheet_pipei.write(count + 1, 23, '1')
                    elif (yuefen_n < 10):
                        Sheet_pipei.write(count + 1, 23, '2')
                    else:
                        Sheet_pipei.write(count + 1, 23, '3')
                    #print ( yuefen_n )

                    flag_title_have_dingliang = data_daochu_deleteuseless[j].Title.find('定量')
                    flag_title_have_shizheng = data_daochu_deleteuseless[j].Title.find('实证')
                    flag_summary_have_dingliang = data_daochu_deleteuseless[j].Summary.find('定量')
                    flag_summary_have_shizheng = data_daochu_deleteuseless[j].Summary.find('实证')

                    if((flag_title_have_dingliang!=-1)or(flag_title_have_shizheng!=-1)or(flag_summary_have_dingliang!=-1)or(flag_summary_have_shizheng!=-1)):
                        Sheet_pipei.write(count + 1, 24, '1')
                    else:
                        Sheet_pipei.write(count + 1, 24, '0')

                    flag_shuangyiliu=0
                    if((data_daochu_deleteuseless[j].Organ.find('编辑部')!=-1)or(data_daochu_deleteuseless[j].Organ.find('杂志社')!=-1)):
                        Sheet_pipei.write(count + 1, 25, '0')
                    elif((data_daochu_deleteuseless[j].Organ.find('学会')!=-1)or(data_daochu_deleteuseless[j].Organ.find('协会')!=-1)):
                        Sheet_pipei.write(count + 1, 25, '1')
                    elif ((data_daochu_deleteuseless[j].Organ.find('集团')!=-1) or (data_daochu_deleteuseless[j].Organ.find('公司')!=-1)):
                        Sheet_pipei.write(count + 1, 25, '2')
                    elif ((data_daochu_deleteuseless[j].Organ.find('部')!=-1) or (data_daochu_deleteuseless[j].Organ.find('署')!=-1)or (data_daochu_deleteuseless[j].Organ.find('国家')!=-1)):
                        Sheet_pipei.write(count + 1, 25, '3')
                    else:
                        for items in shuangyiliu:
                            if(data_daochu_deleteuseless[j].Organ.find(items)!=-1):
                                flag_shuangyiliu =1
                        if(flag_shuangyiliu==1):
                            Sheet_pipei.write(count + 1, 25, '4')
                        else:
                            Sheet_pipei.write(count + 1, 25, '5')

                    if(check_contain_chinese(data_daochu_deleteuseless[j].FirstDuty)):
                        gender_this=ngender.guess(data_daochu_deleteuseless[j].FirstDuty)

                        if(gender_this[0] == 'male'):
                            Sheet_pipei.write(count + 1, 20, '男')
                            Sheet_pipei.write(count + 1, 21,gender_this[1])
                            Sheet_pipei.write(count + 1, 22, '1')
                        elif(gender_this[0] == 'female'):
                            Sheet_pipei.write(count + 1, 20, '女')
                            Sheet_pipei.write(count + 1, 21, gender_this[1])
                            Sheet_pipei.write(count + 1, 22, '0')

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

