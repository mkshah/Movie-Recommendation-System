import os, copy 
import xlrd
import xlwt
import xlutils
from xlrd import open_workbook
from xlwt import Workbook
import networkx as nx
from networkx.algorithms import bipartite

os.chdir('C:\\Users\\Samsung\\Desktop\\BITS PILANI\\YEAR-2 SEM-2\\Projects\\Recommender Systems\\movie recommenders\\movielens dataset\\hx')

avgrat = open_workbook('movies_clustering_coeff.xls')
sh0 = avgrat.sheet_by_index(0)
row0 = sh0.nrows 

rff = open_workbook('randomforest_file.xlsx')
sh1 = rff.sheet_by_index(0)
row1 = sh1.nrows

f = open('workfile12.txt', 'w')

for i in range(row1):
    uid = sh1.cell_value(i,1)
    if(i!=0):
        uidp = sh1.cell_value(i-1,1)
        if(uid==uidp):
            f.write(str(c))
            f.write("\n")
            continue
    for j in range(row0):
        b = sh0.cell_value(j,0)
        c = sh0.cell_value(j,1)
        b1 = b[1:]
        b2 = b1[:-2]
        b3 = int(b2)
        if(uid==b3):
            f.write(str(c))
            f.write("\n")
            break
        
f.close()    