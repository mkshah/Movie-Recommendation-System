import numpy as np
import os, copy 
import xlrd
import xlwt
import xlutils
from xlrd import open_workbook
from xlwt import Workbook

os.chdir('C:\\Users\\Samsung\\Desktop\\BITS PILANI\\YEAR-2 SEM-2\\Projects\\Recommender Systems\\movie recommenders\\movielens dataset\\hx')

ugdcg = open_workbook('scaledf.xls')
sh1 = ugdcg.sheet_by_index(0)
row1 = sh1.nrows

mat = np.zeros((2113,2113))

for i in range(0,2113):
     for j in range(i+1,2113):
          s = 0
          for k in range(0,20):
               s = s + ((sh1.cell_value(i,k))*(sh1.cell_value(j,k)))
          mat[i][j] = s;
          mat[j][i] = s;

f = open('similarity.txt', 'w')
for i in range(0,2113):
     for j in range(0,2113):
          f.write(str(mat[i][j]))
          f.write(" ")
     f.write("\n");     

f.close()