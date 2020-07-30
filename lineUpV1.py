#The program is to line up daily stock quotes from different markets 
#By Lishen Wang and Steed Huang et al from 
#https://www.linkedin.com/company/institute-for-industrial-technology-research
#July 2020 for Global Gold USD & Stocks analysis

import xlrd
import xlwt
import random


data = xlrd.open_workbook("GoldStockPriceInput.xls")

sheet = data.sheet_by_index(0)

date = sheet.col_values(0)


#The prize of gold
data_Gold = sheet.col_values(1)
data_nsdk = sheet.col_values(2)
data_shanghai = sheet.col_values(3)
data_dld = sheet.col_values(4)
data_London = sheet.col_values(5)
data_Tokyo = sheet.col_values(6)

f = xlwt.Workbook()
sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True)



for i in range(0,5495):
    if(data_nsdk[i+1] == ''):
        a= random.randint(4,12)
        print(a)
        total = 0
        k=0
        dash = int(a/2)
        for j in range(0,a):
            if (data_nsdk[i+1-dash] == ''):
                k+=1
                continue
            total+= float(data_nsdk[i+1-dash])
            data_nsdk[i+1] = total/(a-k)
            dash+=1

for i in range(0,5495):
    if(data_shanghai[i+1] == ''):
        a= random.randint(4,12)
        print(a)
        total = 0
        k=0
        dash = int(a/2)
        for j in range(0,a):
            if (data_shanghai[i+1-dash] == ''):
                k+=1
                continue
            total+= float(data_shanghai[i+1-dash])
            data_shanghai[i+1] = total/(a-k)
            dash+=1

for i in range(0,5495):
    if(data_dld[i+1] == ''):
        a= random.randint(4,12)
        print(a)
        total = 0
        k=0
        dash = int(a/2)
        for j in range(0,a):
            if (data_dld[i+1-dash] == ''):
                k+=1
                continue
            total+= float(data_dld[i+1-dash])
            data_dld[i+1] = total/(a-k)
            dash+=1

for i in range(0,5495):
    if(data_London[i+1] == ''):
        a= random.randint(4,12)
        print(a)
        total = 0
        k=0
        dash = int(a/2)
        for j in range(0,a):
            if (data_London[i+1-dash] == ''):
                k+=1
                continue
            total+= float(data_London[i+1-dash])
            data_London[i+1] = total/(a-k)
            dash+=1

for i in range(0,5495):
    if(data_Tokyo[i+1] == ''):
        a= random.randint(4,12)
        print(a)
        total = 0
        k=0
        dash = int(a/2)
        for j in range(0,a):
            if (data_Tokyo[i+1-dash] == ''):
                k+=1
                continue
            total+= float(data_Tokyo[i+1-dash])
            data_Tokyo[i+1] = total/(a-k)
            dash+=1

for i in range(0,5495):
    sheet1.write(i,0,date[i])
    sheet1.write(i,1,data_Gold[i])
    sheet1.write(i,2,data_nsdk[i])
    sheet1.write(i,3,data_shanghai[i])
    sheet1.write(i,4,data_dld[i])
    sheet1.write(i,5,data_London[i])
    sheet1.write(i,6,data_Tokyo[i])

f.save('GoldStockPriceMid.xls')














