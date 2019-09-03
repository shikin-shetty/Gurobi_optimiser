# -*- coding: utf-8 -*-
"""
Created on Mon Apr 22 14:26:24 2019

@author: shikin
"""

from gurobipy import *
import pandas as pd
import numpy as np
from pandas import DataFrame

path = r'C:\Users\Venzire\Desktop\EM3\Dashboard.xlsm'
sheetname ='ForCode'

df1 = pd.read_excel(path,sheetname,header=1,usecols="A",nrows=1)
factories = df1['factories'].tolist() 

df2 = pd.read_excel(path,sheetname,header=1,usecols="B",nrows=3)
breaks = df2['breaks'].tolist()

df3 = pd.read_excel(path,sheetname,header=1,usecols="C",nrows=3)
modes = df3['modes'].tolist()

df4 = pd.read_excel(path,sheetname,header=1,usecols="D",nrows=1)
terminals = df4['terminals'].tolist()

df5 = pd.read_excel(path, sheetname, header=1, index_col=[0,1,2], usecols="P,Q,R,S").to_dict()
cost=df5['cost']

df7 = pd.read_excel(path, sheetname, header=1, index_col=[0,1,2],usecols="P,Q,R,U").to_dict()
transloadcost=df7['transloadcost']

df8 = pd.read_excel(path, sheetname, header=1,index_col=[0],usecols="I,J",nrows=3).to_dict()
modetonnage=df8['modetonnage']

df9 = pd.read_excel(path, sheetname, header=1,index_col=[0,1,2],usecols="P,Q,R,V").to_dict()
flowcapacity=df9['flowcapacity']

df10 = pd.read_excel(path, sheetname, header=1,index_col=[0],usecols="M,N", nrows=3).to_dict()
capacity=df10['capacity']

df11 = pd.read_excel(path, sheetname, header=1,index_col=[0],usecols="A,F", nrows=1).to_dict()
supply=df11['supply']

df12 = pd.read_excel(path, sheetname, header=1,index_col=[0],usecols="D,G", nrows=1).to_dict()
demand=df12['demand']


# Create a new model
m = Model("transport_problem_1")

# Create variables

# flow1 variable defines flow of material from a factory f to intermediary b using mode p
flow1 = {}

for f in factories:
    for b in breaks:
        for p in modes:
            flow1[f,b,p] = m.addVar(name='%s_to_%s_via_%s' % (f,b,p))

# flow2 variable defines flow of material from an intermediary b to terminal s using mode p
flow2 = {}
for b in breaks:
    for s in terminals:
        for p in modes:
            flow2[b,s,p] = m.addVar(name='%s_to_%s_via_%s' % (b,s,p))

# flow3 variable defines flow of material from a factory f to intermediary b
flow3 ={}

for f in factories:
    for b in breaks:
        flow3[f,b] = m.addVar(name='%s_to_%s' % (f,b))

# flow4 variable defines flow of material from an intermediary b to terminal s
flow4 = {}
for b in breaks:
    for s in terminals:
        flow4[b,s] = m.addVar(name='%s_to_%s' % (b,s))

# flow5 variable defines flow of material from a factory f through intermediary b to terminal s
flow5 = {}
for f in factories:
    for b in breaks:
        for s in terminals:
            flow5[f,b,s] = m.addVar(name='%s_to_%s_to_%s' % (f,b,s))
        
# Integrate new variables
m.update()

# Add supply constraints
for f in factories:
    m.addConstr(quicksum((flow3[f,b] for b in breaks)) == supply[f], 'supply_%s' % (f))

# Add flow conservation constraints on factory f to intermediary b on use of different modes
for f in factories:
    for b in breaks:
        m.addConstr(quicksum(flow1[f,b,p] for p in modes) == flow3[f,b])    

# Add flow in and out constraints on intermediary b
for b in breaks:
    m.addConstr(quicksum(flow3[f,b] for f in factories) == quicksum(flow4[b,s] for s in terminals), 'breaks_%s' % (b))

# Add flow conservation constraints on intermediary b to terminal s on use of different modes
for b in breaks:
    for s in terminals:
        m.addConstr(quicksum(flow2[b,s,p] for p in modes) == flow4[b,s])

# Intergate flow variables into one        
for f in factories:
    for b in breaks:
        for s in terminals:
            m.addConstr((flow3[f,b]+flow4[b,s]) == 2*flow5[f,b,s])
    
# Add demand constraints
for s in terminals:
    m.addConstr(quicksum(flow4[b,s] for b in breaks) == demand[s], 'demand_%s' % (s))
    
# Add flowcapacity constraints
# the maximum amount that can be flowed on a network link
# between factory f to intermediary b using a mode p
for f in factories:
    for b in breaks:
        for p in modes:
            m.addConstr(flow1[f,b,p] <= (flowcapacity[f,b,p]))
            
# the maximum amount that can be flowed on a network link
# between intermediary b to terminal s using a mode p
for b in breaks:
    for s in terminals:
        for p in modes:
            m.addConstr(flow2[b,s,p] <= (flowcapacity[b,s,p]))

# Add capacity constraints of intermediary b       
for b in breaks:
    m.addConstr(quicksum(flow3[f,b] for f in factories) <= capacity[b])            


# Optimize the model.
m.setObjective(quicksum(((cost[f,b,p] + transloadcost[f,b,p]) * flow1[f,b,p]) + ((cost[b,s,p] + transloadcost[b,s,p]) * flow2[b,s,p]) 
for f in factories for b in breaks for s in terminals for p in modes), GRB.MINIMIZE)
    
m.optimize()

# Print solution
m.printAttr('X')
m.write('transportation.lp')

flowvar=[]
flowval=[]
if m.status == GRB.OPTIMAL:
    for f in factories:
        for b in breaks:
            for s in terminals:
                if flow5[f,b,s].X>0:
                    flowvar.append(flow5[f,b,s].varName)
                    flowval.append(flow5[f,b,s].X)                    
                    
totalcost = m.ObjVal
totalcost = '${:,.2f}'.format(totalcost)


flowvar = flowvar + ['Total Transportation Cost']
flowval = flowval + [totalcost]

df13 = DataFrame({'flows':flowvar, 
                  'volume':flowval})

print(df13)    
    
path1 = r'C:\Users\shikin\Desktop\EM3\ForDashboard.xlsx'

# write the result to spreadhseet
df13.to_excel(path1,sheet_name='Sheet1',startrow=0,startcol=0)                                           