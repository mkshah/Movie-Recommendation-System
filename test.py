import os, copy 
import xlrd
import xlwt
import xlutils
from xlrd import open_workbook
from xlwt import Workbook
import networkx as nx
from networkx.algorithms import bipartite



os.chdir('C:\\Users\\Samsung\\Desktop\\BITS PILANI\\YEAR-2 SEM-2\\Projects\\Recommender Systems\\movie recommenders\\movielens dataset\\hx')

u_rm = open_workbook('user_ratedmovies.xlsx')
sh0 = u_rm.sheet_by_index(0)
row0 = sh0.nrows - 1

m_ac = open_workbook('movie_actors.xlsx')
sh1 = m_ac.sheet_by_index(0)
row1 = sh1.nrows - 1

m_di = open_workbook('movie_directors.xlsx')
sh2 = m_di.sheet_by_index(0)
row2 = sh2.nrows - 1

m_ge = open_workbook('movie_genres.xlsx')
sh3 = m_ge.sheet_by_index(0)
row3 = sh3.nrows - 1

m_ta = open_workbook('movie_tags.xlsx')
sh4 = m_ta.sheet_by_index(0)
row4 = sh4.nrows - 1

u_tm = open_workbook('user_taggedmovies.xlsx')
sh5 = u_tm.sheet_by_index(0)
row5 = sh5.nrows - 1



u = []
for i in (range(row0)):
    u.append(sh0.cell_value(i+1,0))
i = 0
for x in u:
    x='u'+str(x)
    u[i] = x
    i = i+1
userID = list(set(u))
userID.sort()

m = []
for i in (range(row0)):
    m.append(sh0.cell_value(i+1,1))
i = 0
for x in m:
    x='m'+str(x)
    m[i] = x
    i = i+1
movieID = list(set(m))
movieID.sort()

a = []
for i in (range(row1)):
    a.append(sh1.cell_value(i+1,1))
i = 0
for x in a:
    x='a'+str(x)
    a[i] = x
    i = i+1
actorID = list(set(a))
actorID.sort()

d = []
for i in (range(row2)):
    d.append(sh2.cell_value(i+1,1))
i = 0
for x in d:
    x='d'+str(x)
    d[i] = x
    i = i+1
directorID = list(set(d))
directorID.sort()

g = []
for i in (range(row3)):
    g.append(sh3.cell_value(i+1,1))
i = 0
for x in g:
    x='g'+str(x)
    g[i] = x
    i = i+1
genreID = list(set(g))
genreID.sort()

t = []
for i in (range(row4)):
    t.append(sh4.cell_value(i+1,1))
i = 0
for x in t:
    x='t'+str(x)
    t[i] = x
    i = i+1
tagID = list(set(t))
tagID.sort()



um = nx.Graph()
um.add_nodes_from(userID, bipartite=0)
um.add_nodes_from(movieID, bipartite=1)
for i in range(row0):
    u = sh0.cell_value(i+1,0)
    u = 'u' + str(u) 
    m = sh0.cell_value(i+1,1)
    m = 'm' + str(m)
    r = sh0.cell_value(i+1,2)
    um.add_weighted_edges_from([(u,m,r)])

ma = nx.Graph()
ma.add_nodes_from(actorID, bipartite=0)
ma.add_nodes_from(movieID, bipartite=1)
for i in range(row1):
    m = sh1.cell_value(i+1,0)
    m = 'm' + str(m)
    a = sh1.cell_value(i+1,1)
    a = 'a' + str(a) 
    r = sh1.cell_value(i+1,3)
    ma.add_weighted_edges_from([(a,m,r)])

md = nx.Graph()
md.add_nodes_from(directorID, bipartite=0)
md.add_nodes_from(movieID, bipartite=1)
for i in range(row2):
    m = sh2.cell_value(i+1,0)
    m = 'm' + str(m)
    d = sh2.cell_value(i+1,1)
    d = 'd' + str(d) 
    md.add_edges_from([(d,m)])

mg = nx.Graph()
mg.add_nodes_from(genreID, bipartite=0)
mg.add_nodes_from(movieID, bipartite=1)
for i in range(row3):
    m = sh3.cell_value(i+1,0)
    m = 'm' + str(m)
    g = sh3.cell_value(i+1,1)
    g = 'g' + str(g) 
    mg.add_edges_from([(g,m)])

mt = nx.Graph()
mt.add_nodes_from(tagID, bipartite=0)
mt.add_nodes_from(movieID, bipartite=1)
for i in range(row4):
    m = sh4.cell_value(i+1,0)
    m = 'm' + str(m)
    t = sh4.cell_value(i+1,1)
    t = 't' + str(t) 
    w = sh4.cell_value(i+1,2)
    mt.add_weighted_edges_from([(t,m,w)])

ut = nx.Graph()
ut.add_nodes_from(userID, bipartite=0)
ut.add_nodes_from(movieID, bipartite=1)
e = []
for i in range(row5):
    u = sh5.cell_value(i+1,0)
    u = 'u' + str(u) 
    m = sh5.cell_value(i+1,1)
    m = 'm' + str(m)
    e.append((u,m))
ute = list(set(e))
ut.add_edges_from(ute, tag=[])
for i in range(row5):
    u = sh5.cell_value(i+1,0)
    u = 'u' + str(u)
    m = sh5.cell_value(i+1,1)
    m = 'm' + str(m)
    t = sh5.cell_value(i+1,2)
    temp0 = copy.deepcopy(ut.get_edge_data(u,m))
    temp1 = temp0.get('tag')
    temp1.append(t)
    ut[u][m]['tag'] = temp1
    temp0.clear()


"""
bot, top = bipartite.sets(um)

dc = bipartite.degree_centrality(um, bot);
nd = nx.average_neighbor_degree(um)
cl = bipartite.clustering(um)
pr = nx.pagerank(um)
dg = um.degree();
"""    
    
"""
>>> b1 = Workbook()
>>> s1 = b1.add_sheet('S1')
>>> i = 0
>>> pr1 = pr.items()
>>> pr2 = list(pr1)
>>> pr2.sort()
>>> for k, v in pr2:
...     r = s1.row(i)
...     r.write(0,k)
...     r.write(1,v)
...     i  = i +1
... 
>>> b1.save('pagerank.xls')
"""
    
    
    

    

    


