# -*- coding: utf-8 -*-
"""
Created on Sat May  4 03:49:15 2019

@author:  shikin
"""

from gurobipy import *
import math 
from geopy.distance import geodesic
from itertools import groupby
import pandas as pd
from pandas import DataFrame
from collections import defaultdict

import pandas as pd
data_1= pd.read_excel(r'C:\Users\Venzire\Desktop\EM3\FacilityProject\MasterFacility.xlsm', sheet_name= 'Sheet3')
facilities = pd.DataFrame(data_1, columns = ['Latitude','Longitude'])
print(facilities)
numFacility=len(facilities)

data_3= pd.read_excel(r'C:\Users\Venzire\Desktop\EM3\FacilityProject\MasterFacility.xlsm', sheet_name= 'Sheet3')
name=[]
for i in data_3["Port Name"]:
    name.append(i)
print(name)    

data_2 = pd.read_excel(r'C:\Users\Venzire\Desktop\EM3\FacilityProject\MasterFacility.xlsm', sheet_name='Sheet4')
clients = pd.DataFrame(data_2, columns = ['Latitude','Longitude'])
numClient=len(clients)
print(numClient)
print(clients.loc[0])
print(geodesic(clients.loc[0],facilities.loc[3]).miles)

d = {}

for i in range(numClient):
  for j in range(numFacility):
      d[(i,j)] = geodesic(clients.loc[i], facilities.loc[j]).miles

client=[]   
for i in range(numClient):
  for j in range(numFacility):
      if d[(i,j)]<=50:
          client.append([(i,j)])

lis1=[]
for i in range(len(client)):
    lis1.append(client[i][0][0])
    
lis2=[]
for i in range(len(client)):
    lis2.append(client[i][0][1]) 
    
count=[len(list(group)) for key, group in groupby(lis1)]    

lis3 = defaultdict(list)
for k, v in zip(lis1,lis2):
    lis3[k].append(v)

lis3 = dict(lis3)
print(lis3)

clientlist = list(lis3.keys())
facilitylist = list(set(lis2))

m = Model()

# Add variables

consider = {}
isopen = {}

path= r'C:\Users\Venzire\Desktop\EM3\FacilityProject\MasterFacility.xlsm'
sheet_name = 'Sheet2'
df1 = pd.read_excel(path,sheet_name,header=0,usecols="A",nrows=1)
df2 = pd.read_excel(path,sheet_name,header=0,usecols="B",nrows=1)
k = df1['FacilitiesNo']
dist = df2['Distance']
print(dist)

for i in clientlist:
          consider[i] = m.addVar(vtype='B', name= "clientconsider%s" %i)

for j in facilitylist:
        isopen[j] = m.addVar(vtype='B', name="facilityopen%s" %j)
    
#for i in range(numClients):
#    consider[i] = m.addVar(vtype='B', name="clientopen%s" %i)    


d = {}  # Distance matrix (not a variable)
for i in clientlist:
    for j in lis3.get(i):
        d[(i,j)] = geodesic(clients.loc[i], facilities.loc[j]).miles
        
m.update()

# Add constraints

for i in clientlist:
       m.addConstr(quicksum(isopen[j] for j in lis3.get(i))>=consider[i])

m.addConstr(quicksum(isopen[j] for j in facilitylist) == k)    

        
m.setObjective(quicksum(consider[i] for i in clientlist),GRB.MAXIMIZE)

m.optimize()

for v in m.getVars():
        print('%s %g' % (v.varName, v.x))

m.printAttr('X')
m.write('facilityPY1.lp')


if m.status == GRB.OPTIMAL:
    indices = [j for j in facilitylist if isopen[j].X > 0]

print(indices)

for j in indices:
    print(facilities.loc[j])

terminals_loc=[]
for j in indices:
    terminals_loc.append((name[j]))
print(terminals_loc)

a = m.ObjVal/numClient

a = [a] + terminals_loc

table={'Distribution_Center': a}    
dfr = pd.DataFrame(table)
path = r'C:\Users\shikin\Desktop\EM3\FacilityProject\Forcode.xlsx'
dfr.to_excel(path,sheet_name='Sheet1',startrow=0,startcol=0)        