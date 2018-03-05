import copy
from xlrd import open_workbook
from xlwt import *
import heapq
import os

os.chdir('C:\\Users\\Samsung\\Desktop\\BITS PILANI\\YEAR-2 SEM-2\\Projects\\Recommender Systems\\movie recommenders\\movielens dataset\\hx')
    
rb = open_workbook('Similarity.xlsx')
r_sheet = rb.sheet_by_index(0)


heap = []
def make_queue(user):
    for col in range(0,r_sheet.ncols):
        val = (r_sheet.cell(user,col).value)
        if val:         # not going in for 0 
            num = float(val)
            #print "received value",num,"\n"
            
            if (col > 10):
             #   print "min is ", heap[0][0], " ", heap[0][1],"\t" 
                if(num > heap[0][0]):
                   replaced = heapq.heapreplace(heap, (num, col))
              #     print "replaced ", replaced[0], " with ",num
            else:
                    heapq.heappush(heap, (num, col))
               #     print "pushed ", col , "\n"


make_queue(0)       # call for some user_number here, users start from 0, so this makes queue for 1st row in excel sheet

ordered = []
while heap:
    ordered.append(heapq.heappop(heap))

ordered.sort(key=lambda x: x[0], reverse = True)

print ("User Number","\t \t","Similarity")

for x in ordered:
    print (x[1],"\t", x[0],"\n")
