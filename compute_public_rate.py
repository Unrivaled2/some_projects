#!/usr/bin/env python
# coding: utf-8

# In[72]:


import xlrd
import xlwt
from xlutils.copy import copy


# In[127]:


###读取三张表写成字典，key：省，城市，年，月，value：公开率
refer = xlrd.open_workbook(r'/home/cug/data/reference.xlsx')
refer_sheet1 = refer.sheet_by_index(0)
refer_sheet2 = refer.sheet_by_index(1)
refer_sheet3 = refer.sheet_by_index(2)


# In[155]:


##读取信息到refer_sum,写入sum_sheet
refer_sum = {}
rb = xlrd.open_workbook('/home/cug/data/summary.xlsx') 
wb = copy(rb)
sum_sheet = wb.get_sheet(3) 
sum_sheet_r = rb.sheet_by_index(3)


# In[129]:


sum_sheet.write(2,3,1)
wb.save('/home/cug/data/summary.xlsx')


# In[156]:


def read_refer1():
    global refer_sum
    global sum_sheet
    global refer_sheet1
    global refer_sheet2
    global refer_sheet3
    global wb
    for i in range(2734):
        if i==0:
            continue
        province = refer_sheet1.cell(i,5).value
        area = refer_sheet1.cell(i,6).value
        year = refer_sheet1.cell(i,1).value
        public = False
        if (refer_sheet1.cell(i,14).value=='是'):
            public = True
        if (refer_sheet1.cell(i,14).value=='是/否'):
            if (refer_sheet1.cell(i,15).value=='是'):
                public = True
                print('found true/false')
        month = refer_sheet1.cell(i,2).value
        key = (province,area,year,month)
        if(key in refer_sum.keys()) & public:                       ##如果键值存在，value加1，都则update添加
            value1 = refer_sum[key][0] + 1
            value2 = refer_sum[key][1] + 1
            refer_sum.update({key:(value1,value2)})
        if (key in refer_sum.keys()) & (not public):
            value1 = refer_sum[key][0]
            value2 = refer_sum[key][1] + 1
            refer_sum.update({key:(value1,value2)})
        if (not key in refer_sum.keys()) & public:
            refer_sum.update({key:(1,1)})
        if (not key in refer_sum.keys()) & (not public):
            refer_sum.update({key:(0,1)})
        
    
def read_refer2():
    count = 0
    global refer_sum
    global sum_sheet
    global refer_sheet1
    global refer_sheet2
    global refer_sheet3
    global wb
    for i in range(31444):
        if i==0:
            continue
        province = refer_sheet2.cell(i,5).value
        area = refer_sheet2.cell(i,6).value
        year = refer_sheet2.cell(i,1).value
        public = False
        if (refer_sheet2.cell(i,14).value=='是'):
            public = True
        if (refer_sheet2.cell(i,14).value=='是/否'):
            if (refer_sheet2.cell(i,15).value=='是'):
                public = True
                
        month = refer_sheet2.cell(i,2).value
        key = (province,area,year,month)
        if(key in refer_sum.keys()) & public:                       ##如果键值存在，value加1，都则update添加
            value1 = refer_sum[key][0] + 1
            value2 = refer_sum[key][1] + 1
            refer_sum.update({key:(value1,value2)})

        if (key in refer_sum.keys()) & (not public):
            value1 = refer_sum[key][0]
            value2 = refer_sum[key][1] + 1
            refer_sum.update({key:(value1,value2)})
            #print('2')
        if (not key in refer_sum.keys()) & public:
            refer_sum.update({key:(1,1)})
            #print('3')
        if (not key in refer_sum.keys()) & (not public):
            refer_sum.update({key:(0,1)})
            #print('4')
        count += 1
        print('count is %d'%count)
        
def read_refer3():
    count = 0
    global refer_sum
    global sum_sheet
    global refer_sheet1
    global refer_sheet2
    global refer_sheet3
    global wb
    for i in range(16585):
        if i==0:
            continue
        province = refer_sheet3.cell(i,5).value
        area = refer_sheet3.cell(i,6).value
        year = refer_sheet3.cell(i,1).value
        public = False
        if (refer_sheet3.cell(i,14).value=='是'):
            public = True
        if (refer_sheet3.cell(i,14).value=='是/否'):
            if (refer_sheet3.cell(i,15).value=='是'):
                public = True
                print('found true/false')
        month = refer_sheet3.cell(i,2).value
        key = (province,area,year,month)
        if(key in refer_sum.keys()) & public:                       ##如果键值存在，value加1，都则update添加
            value1 = refer_sum[key][0] + 1.0
            value2 = refer_sum[key][1] + 1.0
            refer_sum.update({key:(value1,value2)})
        if (key in refer_sum.keys()) & (not public):
            value1 = refer_sum[key][0]
            value2 = refer_sum[key][1] + 1.0
            refer_sum.update({key:(value1,value2)})
        if (not key in refer_sum.keys()) & public:
            refer_sum.update({key:(1.0,1.0)})
        if (not key in refer_sum.keys()) & (not public):
            refer_sum.update({key:(0.0,1.0)})
        count += 1
        print('count is %d'%count)


# In[157]:


read_refer1()
read_refer2()
read_refer3()


# In[172]:


def ensure_write():
    global sb
    global wb
    global sum_sheet
    global sum_sheet_r
    wb.save('/home/cug/data/summary.xlsx')
    print('ensure')
    rb = xlrd.open_workbook('/home/cug/data/summary.xlsx') 
    wb = copy(rb)
    sum_sheet = wb.get_sheet(3) 
    sum_sheet_r = rb.sheet_by_index(3)
    

i = 2


def write():
    global i
    global refer_sum
    global sum_sheet
    global refer_sheet1
    global refer_sheet2
    global refer_sheet3
    global sum_sheet_r
    for key,value in refer_sum.items():
        month = key[3]
        province = key[0]
        distinct = key[1]
        year = key[2]
        tag = False
        pos = 0
        rb = xlrd.open_workbook('/home/cug/data/summary.xlsx') 
        wb = copy(rb)
        sum_sheet = wb.get_sheet(3) 
        sum_sheet_r = rb.sheet_by_index(3)
        if month=='':
            print('月份为空')
            continue
        if(month=='1月'):
            month=1
        if(month=='3月'):
            month=3
        if(month=='7月'):
            month=7
        if(month=='8月'):
            month=8
        if(month=='9月'):
            month=9
        if(month=='10月'):
            month=10
        if(month=='12月'):
            month=12
        if(month=='4月'):
            month=4
        if(month=='5月'):
            month=5
        if(month=='6月'):
            month=6
        if(month=='2月'):
            month=2
        if(month=='11月'):
            month=11
        result = value[0]/value[1]
        for j in range(i):
            if j==0:
                continue
            if j==1:
                continue
            if(sum_sheet_r.cell(j,0).value==key[0]):
                if(sum_sheet_r.cell(j,1).value==key[1]):
                    pos = j
                    tag = True
                    if(year==2015.0):
                        sum_sheet.write(pos,int(month+1),result)
                        wb.save('/home/cug/data/summary.xlsx')
                        break
                    if(year==2016.0):
                        sum_sheet.write(pos,int(month+13),result)
                        wb.save('/home/cug/data/summary.xlsx')
                        break
                    if(year==2017.0):
                        sum_sheet.write(pos,int(month+25),result)
                        wb.save('/home/cug/data/summary.xlsx')
                        break
        if(tag):
            continue
        if(year==2015.0):
            rb = xlrd.open_workbook('/home/cug/data/summary.xlsx') 
            wb = copy(rb)
            sum_sheet = wb.get_sheet(3)
            
            if(tag):
                sum_sheet.write(pos,0,province)
                sum_sheet.write(pos,1,distinct)
                sum_sheet.write(pos,int(month+1),result)
                wb.save('/home/cug/data/summary.xlsx')
            else:
                sum_sheet.write(i,0,province)
                sum_sheet.write(i,1,distinct)
                sum_sheet.write(i,int(month+1),result)
                wb.save('/home/cug/data/summary.xlsx')
                i = i+1
                print('2015')
        if(year==2016.0):
            rb = xlrd.open_workbook('/home/cug/data/summary.xlsx') 
            wb = copy(rb)
            sum_sheet = wb.get_sheet(3)
            if(tag):
                sum_sheet.write(pos,0,province)
                sum_sheet.write(pos,1,distinct)
                sum_sheet.write(pos,int(month+13),result)
                wb.save('/home/cug/data/summary.xlsx')
            else:
                sum_sheet.write(i,0,province)
                sum_sheet.write(i,1,distinct)
                sum_sheet.write(i,int(month+13),result)
                wb.save('/home/cug/data/summary.xlsx')
                i = i+1
                print('2016')
        if(year==2017.0):
            rb = xlrd.open_workbook('/home/cug/data/summary.xlsx') 
            wb = copy(rb)
            sum_sheet = wb.get_sheet(3)
            if(tag):
                sum_sheet.write(pos,0,province)
                sum_sheet.write(pos,1,distinct)
                sum_sheet.write(pos,int(month+25),result)
                wb.save('/home/cug/data/summary.xlsx')
            else:
                sum_sheet.write(i,0,province)
                sum_sheet.write(i,1,distinct)
                sum_sheet.write(i,int(month+25),result)
                wb.save('/home/cug/data/summary.xlsx')
                i = i+1
                print('2017')
        
        

write()


# In[174]:


for key,value in refer_sum.items():
    if(key[3]==''):
        print(key)
        print(value)


# In[ ]:



