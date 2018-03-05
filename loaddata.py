import os, copy 
import numpy as np
import xlrd
import xlwt
import xlutils
from xlrd import open_workbook
from xlwt import Workbook

os.chdir('C:\\Users\\Samsung\\Desktop\\BITS PILANI\\YEAR-2 SEM-2\\Projects\\Recommender Systems\\movie recommenders\\movielens dataset\\hx')

pdat = np.zeros((205598))

data = open_workbook('predicted_RFC_n50.xlsx')
sh0 = data.sheet_by_index(0)
row0 = sh0.nrows - 1

for i in range(205598):
    pdat[i] = sh0.cell_value(i,0)