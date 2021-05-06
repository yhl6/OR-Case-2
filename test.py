from gurobipy import *
a = [1, 2, 3]
b = [4, 5, 6]
model = Model('moder')
x = model.addVars(2, 3, lb=2, vtype=GRB.INTEGER)
y = model.addVars(2, 3, lb=0, vtype=GRB.INTEGER)

model.setObjective(quicksum(x.select()), GRB.MINIMIZE)

model.addConstr(quicksum(x.select('*'                                                                                                                                                                                             , i) for i in range(3)))

model.optimize()

