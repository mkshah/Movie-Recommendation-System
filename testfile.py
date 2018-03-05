import os, copy 
import xlrd
import xlwt
import xlutils
from xlrd import open_workbook
from xlwt import Workbook
import networkx as nx
from networkx.algorithms import bipartite

os.chdir('C:\\Users\\Samsung\\Desktop\\BITS PILANI\\YEAR-2 SEM-2\\Projects\\Recommender Systems\\movie recommenders\\movielens dataset\\hx')
f = open('predicted.txt', 'w')

for i in range(output.size):
    f.write(str(output[i]))
    f.write("\n")

f.close()

