from xlutils.copy import copy
from xlrd import open_workbook
from xlwt import *
import operator
rb = open_workbook('F:/workableDATASETnew.xls', formatting_info = True)

wb = Workbook()
Sheet1 = wb.add_sheet('Sheet1')
write_row = 0

r_sheet = rb.sheet_by_index(0)

dict = {}           #to track user id which are seen already
genres  = {}
userid = 0
profileDict = {}
global NumMovies     #for number of movies seen by user

def calcUserProf(id,NumMovies):
    global write_row
    profileDict.clear()
              
    for key in genres.keys():
        sum = 0
        list = genres.get(key)
        for rating in list:
            sum+=rating    
        profileDict[key] = sum
    
    print "count = ",NumMovies
    for key in profileDict:         #first store sum of ratings of each genre then divide by total number of ratings
        profileDict[key] = profileDict.get(key)/NumMovies
    Top3 = sorted(profileDict.iteritems(), key=operator.itemgetter(1), reverse=True)[:3]
    print "profile for:", id
    Sheet1.write(write_row,0,id)
    write_row =  write_row +1
    for key,value in Top3:
        Sheet1.write(write_row,0,key)
        write_row =  write_row +1
        print "genre =",key,"average rating = ",value
   # print "profile for user",id, "is" , profileDict
        
for row in range(1,r_sheet.nrows):
    prevUserId = userid
    userid = (int)(r_sheet.cell(row,0).value)
    
    if dict.get(userid):            #if this user id is already seen
        NumMovies = NumMovies+1 
        list = r_sheet.cell(row,3).value  #stores list of genres
        list2 = list.split(',')
        for element in list2:
            genres.setdefault(element, []).append(r_sheet.cell(row,2).value)         #append the rating IN FLOAT to the list with key as genre name
            
    
    else:
        dict.setdefault(userid, 'true')
        if prevUserId!=0:
            calcUserProf(prevUserId,NumMovies)         #for current genre dictionary
        NumMovies=1
        genres.clear()                                  # for the current row, add its genres
        list = r_sheet.cell(row,3).value  #stores list of genres
        list2 = list.split(',')
        for element in list2:
            genres.setdefault(element, []).append(r_sheet.cell(row,2).value)    
       

calcUserProf(prevUserId,NumMovies)        #for last user
print len(dict)         #to see number of users successfully processed
#print genres
wb.save('F:/Profiles.xls')


        
    
