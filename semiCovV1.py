#The program is to line up daily stock quotes from different markets 
#By Lishen Wang and Steed Huang et al from 
#https://www.linkedin.com/company/institute-for-industrial-technology-research
#August 2020 for Global Gold USD & Stocks analysis including holidays

import xlwt
import xlrd
import math

def get_average(data_Gold):
    length = 5495
    avg_Gold = []
    for i in range(1,length): 
        total = 0
        if(i >(length-14)):
            for k in range(0,length-i):
                total += data_Gold[i+k]
            avg_Gold.append(total/(length-i))
        else:  
            for j in range(0,14):
                total += data_Gold[i+j] 
            avg_Gold.append(total/14)

    return avg_Gold

data = xlrd.open_workbook("GoldStockPriceMid.xls")
sheet = data.sheet_by_index(0)
date = sheet.col_values(0)
f = xlwt.Workbook()
sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True) 

#The prize of gold
data_Gold = sheet.col_values(1)
data_nsdk = sheet.col_values(2)
data_shanghai = sheet.col_values(3)
data_dld = sheet.col_values(4)
data_London = sheet.col_values(5)
data_Tokyo = sheet.col_values(6)

avg_Gold = []
avg_nsdk = []
avg_shanghai = []
avg_dld = []
avg_London = []
avg_Tokyo = []

name_shares = sheet.row_values(0)
avg_Gold = get_average(data_Gold)
avg_nsdk = get_average(data_nsdk)
avg_shanghai = get_average(data_shanghai)
avg_dld = get_average(data_dld)
avg_London = get_average(data_London)
avg_Tokyo = get_average(data_Tokyo)

real = []
for i in range(1,5495):
    a = (float(data_Gold[i])-float(avg_Gold[i-1]))*(float(data_nsdk[i])-float(avg_nsdk[i-1]))*(float(data_shanghai[i])-float(avg_shanghai[i-1]))*(float(data_dld[i])-float(avg_dld[i-1]))*(float(data_London[i])-float(avg_London[i-1]))*(float(data_Tokyo[i])-float(avg_Tokyo[i-1]))
    if(a>0):
        real.append(pow(a,1/6))
    else:
        real.append(0)

img = []
for i in range(1,5495):
    b = -(float(data_Gold[i])-float(avg_Gold[i-1]))*(float(data_nsdk[i])-float(avg_nsdk[i-1]))*(float(data_shanghai[i])-float(avg_shanghai[i-1]))*(float(data_dld[i])-float(avg_dld[i-1]))*(float(data_London[i])-float(avg_London[i-1]))*(float(data_Tokyo[i])-float(avg_Tokyo[i-1]))
    if(b>0):
        img.append(-pow(b,1/6))
    else:
        img.append(0)



for i in range(1,5495):
    sheet1.write(i,0,date[i])
    sheet1.write(i,1,data_Gold[i])
    sheet1.write(i,2,avg_Gold[i-1])
    sheet1.write(i,3,data_nsdk[i])
    sheet1.write(i,4,avg_nsdk[i-1])
    sheet1.write(i,5,data_shanghai[i])
    sheet1.write(i,6,avg_shanghai[i-1])
    sheet1.write(i,7,data_dld[i])
    sheet1.write(i,8,avg_dld[i-1])
    sheet1.write(i,9,data_London[i])
    sheet1.write(i,10,avg_London[i-1])
    sheet1.write(i,11,data_Tokyo[i])
    sheet1.write(i,12,avg_Tokyo[i-1])
    sheet1.write(i,13,real[i-1])
    sheet1.write(i,14,img[i-1])
f.save('GoldStockPriceOutput.xls')