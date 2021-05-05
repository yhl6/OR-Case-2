# import gurobipy as gp
# from gurobipy import GRB

# commodities = ['Pencils', 'Pens']
# nodes = ['Detroit', 'Denver', 'Boston', 'New York', 'Seattle']

# arcs, capacity = gp.multidict({
#     ('Detroit', 'Boston'):   100,
#     ('Detroit', 'New York'):  80,
#     ('Detroit', 'Seattle'):  120,
#     ('Denver',  'Boston'):   120,
#     ('Denver',  'New York'): 120,
#     ('Denver',  'Seattle'):  120})
# print(arcs)
# print(capacity)
import numpy as np
import pandas as pd
from gurobipy import *

# df = pd.DataFrame(np.random.randint(-100,100,size=(100, 4)), columns=list('ABCD'))
# print(df.head())
# a = Model('a')

# b = a.addVars(4, 100, lb=0, vtype=GRB.INTEGER, name='b')

# a.setObjective(b.prod(df), GRB.MAXIMIZE)
a = [*range(1, 10)]

print([i - 1 for i in a])

 
df=pd.DataFrame(np.arange(12).reshape((3,4)),index=['one','two','thr'],columns=list('abcd'))
print(df)