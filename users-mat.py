import numpy as np
import os, copy 
import xlrd
import xlwt
import xlutils
from xlrd import open_workbook
from xlwt import Workbook

os.chdir('C:\\Users\\Samsung\\Desktop\\BITS PILANI\\YEAR-2 SEM-2\\Projects\\Recommender Systems\\movie recommenders\\movielens dataset\\hx')

ug = open_workbook('Allusers1.xlsx')
sh0 = ug.sheet_by_index(0)
row0 = sh0.nrows

nm = open_workbook('noofmov.xlsx')
sh1 = nm.sheet_by_index(0)
row1 = sh1.nrows

b1 = Workbook()
s1 = b1.add_sheet('S1')
    
for i in range(row0):
    for j in range(20):
        rat = sh0.cell_value(i,j)
        avg = sh0.cell_value(i,20)
        sd = sh0.cell_value(i,21)
        val1 = (rat-avg)/sd
        num = sh1.cell_value(i,j)
        tnum = sh1.cell_value(i,20)
        val2 = num/tnum
        val = val1*pow(1.25,val2)
        if(val1<0):
            dif = val1 - val
            val = val + (2*dif)            
        s1.write(i,j,val)

b1.save('scaledf.xls')