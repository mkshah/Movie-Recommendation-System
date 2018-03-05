import numpy as np
import os, copy 
import xlrd
import xlwt
import xlutils
from xlrd import open_workbook
from xlwt import Workbook
import networkx as nx
from networkx.algorithms import bipartite

os.chdir('C:\\Users\\Samsung\\Desktop\\BITS PILANI\\YEAR-2 SEM-2\\Projects\\Recommender Systems\\movie recommenders\\movielens dataset\\hx')

rfr = open_workbook('random_forest_randomized.xlsx')
sh0 = rfr.sheet_by_index(0)
row0 = sh0.nrows

data = np.zeros((855598,15))

for i in range(855598):
    for j in range(15):
        data[i][j] = sh0.cell_value(i,j)
        
                
"""
from sklearn.ensemble import RandomForestClassifier 
forest = RandomForestClassifier(n_estimators = 100)
forest = forest.fit(data[0:650000,0:13],data[0:650000,14])
output = forest.predict(data[650000:855598,0:13])
"""        
        
        