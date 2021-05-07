# from gurobipy import *
# a = [1, 2, 3]
# b = [4, 5, 6]
# model = Model('moder')
# x = model.addVars(2, 3, lb=2, vtype=GRB.INTEGER)
# y = model.addVars(2, 3, lb=0, vtype=GRB.INTEGER)
#
# model.setObjective(quicksum(x.select()), GRB.MINIMIZE)
#
# model.addConstr(quicksum(x.select('*'                                                                                                                                                                                             , i) for i in range(3)))
#
# model.optimize()

import pandas as pd
import numpy as np

def add(x, y):
    print(x + y - 2)

print('hi')

# pd.set_option('display.max_columns', None)
# arr = [['a', 'a', 'a', 'b', 'b', 'b', 'c', 'c', 'c'], ['E', 'A', 'O'] * 3]

# tuples = list(zip(*arr))
#
# index = pd.MultiIndex.from_tuples(tuples, names=['first', 'second'])

# df = pd.DataFrame([[*range(1, 10)], [*range(4, 13)], [*range(10, 19)]], index=['A', 'B', 'C'], columns=index)
# print(df)
# print(df.loc['A', ('a', 'O')])
