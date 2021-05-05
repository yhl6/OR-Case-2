from gurobipy import *
import numpy as np
import pandas as pd
from pathlib import Path

# def find_vendor_price(Vendor_product, Vendor_cost):
#
'''
讀取資料
'''
DATA_PATH = Path(__file__).resolve().parent / 'data'
file_name = DATA_PATH / 'OR109-2_case02_data_s1.xlsx'
df = pd.ExcelFile(file_name)
Demand = pd.read_excel(df, sheet_name='Demand', index_col=0)

# 期初存貨
Initial_inventory = pd.read_excel(df, sheet_name='Initial inventory', index_col=0)
Initial_inv = Initial_inventory['Initial inventory']

# 變動運輸成本
Shipping_cost = pd.read_excel(df, sheet_name='Shipping cost', index_col=0)
Express_delivery = Shipping_cost['Express delivery']
Air_freight = Shipping_cost['Air freight']

# 在途存貨
In_transit = pd.read_excel(df, sheet_name='In-transit', index_col=0)

# 產品體積
Size = pd.read_excel(df, sheet_name='Size', index_col=0)
Cubic_meter = Size['Size']

# 售價、購買成本、持有成本
Price_and_cost = pd.read_excel(df, sheet_name='Price and cost', index_col=0)
Sales = Price_and_cost['Sales price']
Purchasing_cost = Price_and_cost['Purchasing cost']
Holding_cost = Price_and_cost['Holding cost']

# 銷售順延成本、銷售損失成本、銷售順延比率
Shortage = pd.read_excel(df, sheet_name='Shortage', index_col=0)
Backorder = Shortage['Backorder']  # 0.05 p_i
Lost_sales = Shortage['Lost sales']  # 機會成本
Backorder_percentage = Shortage['Backorder percentage']  # 願意繼續等的機率

# 其他資料
Vendor_product = pd.read_excel(df, sheet_name='Vendor-Product', index_col=0)
Vendor_cost = pd.read_excel(df, sheet_name='Vendor cost', index_col=0)

Bounds = pd.read_excel(df, sheet_name='Bounds', index_col=0)
Lower_bound = Bounds['Minimum order quantity (if an order is placed)']
Conflict = pd.read_excel(df, sheet_name='Conflict', index_col=0)
Conflict_tl = tuplelist(zip(Conflict['Product 1'], Conflict['Product 2']))

Fixed_cost = [100, 80, 50]

N, M = Demand.shape  # 產品數、月份數
R = Vendor_cost.shape[0]  # 供應商的數量
J = Conflict.shape[0]  # 衝突的數量(共有幾種衝突)

ProductID = range(1, N + 1)
MonthID = range(1, M + 1)
Shipping_method = range(1, 4)
VendorID = range(1, R + 1)
ConflictID = range(1, J + 1)

# 模型的部分
p2 = Model("p2")

# Non-binary Variables
x = p2.addVars(ProductID, Shipping_method, MonthID, lb=0, vtype=GRB.INTEGER, name='x')  # x[i, k, t]
y = p2.addVars(ProductID, MonthID, lb=-1000000000000, vtype=GRB.INTEGER, name='y')  # y 期末存貨 (現在有分正負)
g = p2.addVars(MonthID, lb=0, vtype=GRB.INTEGER, name='g')  # 高斯符號
yp = p2.addVars(ProductID, MonthID, lb=0, vtype=GRB.INTEGER, name='yp')  # 期末存貨為正時的值
yn = p2.addVars(ProductID, MonthID, lb=0, vtype=GRB.INTEGER, name='yn')  # 期末存貨為負時(未滿足需求)時的值

# Binary Variables
z = p2.addVars(Shipping_method, MonthID, lb=0, vtype=GRB.BINARY, name='z')  # z[k, t] = 1 if method k is used in month t
d = p2.addVars(ProductID, MonthID, lb=0, vtype=GRB.BINARY,
               name='d')  # d[i, t] = 1 if the amount of y[i, t] is positive or = 0 otherwise.
a = p2.addVars(ProductID, MonthID, lb=0, vtype=GRB.BINARY,
               name='a')  # a[i, t] = 1 if product i is ordered in month t or = 0 otherwise.
o = p2.addVars(VendorID, lb=0, vtype=GRB.BINARY, name='o')  # o[i, t] = 1 if we order any product from vendor r

# 更新變數
p2.update()

# 目標變數
# p2.setObjective(
#     quicksum(Express_delivery[i] * quicksum(x[i, 1, t] for t in MonthID) for i in ProductID)  # 變動成本
#     + quicksum(Air_freight[i] * quicksum(x[i, 2, t] for t in MonthID) for i in ProductID)
#     + quicksum(Fixed_cost[k - 1] * quicksum(z[k, t] for t in MonthID) for k in Shipping_method)  # 固定成本
#     + quicksum(
#         Purchasing_cost[i] * quicksum(x[i, k, t] for k in Shipping_method for t in MonthID) for i in ProductID)  # 購買成本
#     + quicksum(Holding_cost[i] * quicksum(yp[i, t] * d[i, t] for t in MonthID) for i in ProductID)
#     + quicksum(g[t] * 2750 for t in MonthID)
#     + quicksum(
#         Lost_sales[i] * quicksum(yn[i, t] * (1 - d[i, t]) * (1 - Backorder_percentage[i]) for t in MonthID) for i in
#         ProductID)
#     + quicksum(
#         Backorder[i] * quicksum(yn[i, t] * (1 - d[i, t]) * Backorder_percentage[i] for t in MonthID) for i in ProductID)
#     + quicksum(Vendor_cost.loc[r, 'Ordering cost'] * o[r] for r in VendorID), GRB.MINIMIZE)

p2.setObjective(
    quicksum(Express_delivery[i] * quicksum(x.select('*', 1, '*')) for i in ProductID)  # 變動成本
    + quicksum(Air_freight[i] * quicksum(x.select('*', 2, '*')) for i in ProductID)
    + quicksum(Fixed_cost[k - 1] * quicksum(z.select(k, '*')) for k in Shipping_method)  # 固定成本
    + quicksum(
        Purchasing_cost[i] * quicksum(x.select(i, '*', '*')) for i in ProductID)  # 購買成本
    + quicksum(Holding_cost[i] * quicksum(yp[i, t] * d[i, t] for t in MonthID) for i in ProductID)
    + quicksum(g.select() * 2750)
    + quicksum(
        Lost_sales[i] * quicksum(yn[i, t] * (1 - d[i, t]) * (1 - Backorder_percentage[i]) for t in MonthID) for i in
        ProductID)
    + quicksum(
        Backorder[i] * quicksum(yn[i, t] * (1 - d[i, t]) * Backorder_percentage[i] for t in MonthID) for i in ProductID)
    + quicksum(Vendor_cost.loc[r, 'Ordering cost'] * o[r] for r in VendorID), GRB.MINIMIZE)

# 限制式

# y[i, t] Constraint
p2.addConstrs((y[i, 1] == Initial_inv[i] + In_transit.loc[i, 1] - Demand.loc[i, 1] for i in ProductID), 'Week 1')
p2.addConstrs((y[i, 2] == yp[i, 1] * d[i, 1] - yn[i, 1] * Backorder_percentage[i] + In_transit.loc[i, 2] + x[i, 1, 1] -
               Demand.loc[i, 2] for i in ProductID), 'Week 2')
p2.addConstrs((y[i, 3] == yp[i, 2] * d[i, 2] - yn[i, 2] * Backorder_percentage[i] + In_transit.loc[i, 3] + x[i, 1, 2] +
               x[i, 2, 1] - Demand.loc[i, 3] for i in ProductID), 'Week 3')
p2.addConstrs((y[i, t] == yp[i, t - 1] * d[i, t - 1] - yn[i, t - 1] * (1 - d[i, t - 1]) * Backorder_percentage[
    i] + quicksum(x[i, k, t - k] for k in Shipping_method) - Demand.loc[i, t]
               for i in ProductID for t in MonthID[3:]))

# Z Constraint
for t in MonthID:
    for k in Shipping_method:
        p2.addConstr(quicksum(x[i, k, t] for i in ProductID) <= (
                quicksum(Demand.loc[i, t] for i in ProductID for t in MonthID) - quicksum(
            Initial_inv[i] for i in ProductID)) * z[k, t], 'Z constraint')

# Y的binary constraint
p2.addConstrs(d[i, t] <= yp[i, t] for i in ProductID for t in MonthID)

# inventory constraint (yp >= 0? yn >= 0?)
p2.addConstrs(yp[i, t] >= y[i, t] for i in ProductID for t in MonthID)
p2.addConstrs(yp[i, t] <= y[i, t] * d[i, t] for i in ProductID for t in MonthID)
p2.addConstrs(yn[i, t] >= -y[i, t] for i in ProductID for t in MonthID)
p2.addConstrs(yn[i, t] <= -y[i, t] * (1 - d[i, t]) for i in ProductID for t in MonthID)

# ceiling constraint
p2.addConstrs(g[t] >= quicksum(Cubic_meter[i] * x[i, 3, t] for i in ProductID) / 30 for t in MonthID)
p2.addConstrs(
    g[t] <= quicksum(Cubic_meter[i] * x[i, 3, t] for i in ProductID) / 30 + 0.999999999999999999 for t in MonthID)

# conflict constraint
p2.addConstrs((a[i, t] + a[j, t] <= 1) for i, j in Conflict_tl for t in MonthID)

# Lower Bound constraint
p2.addConstrs(quicksum(x.select(i, '*', t)) >= Lower_bound[i] * a[i, t] for i in ProductID for t in MonthID)
p2.addConstrs(quicksum(x.select(i, '*', t)) <= quicksum(Demand.loc[i]) for i in ProductID for t in MonthID)


'''
p2.addConstrs((y[i, 1] == Initial_Inventory[i] - Demandlist[i, 1] for i in ProductID), 'March')
p2.addConstrs((y[i, 2] == yp[i, 1]*d[i, 1] - yn[i, 1]*Backorder_percentage[i] + x[i, 1][0] - Demandlist[i, 2] + April_intransit[i] for i in ProductID), 'April')
p2.addConstrs((y[i][2] == yp[i, 2]*d[i, 2] - yn[i, 2]*Backorder_percentage[i] + x[i, 1][1] + x[i, 2][0] - Demandlist[i][2] + May_intransit[i] for i in ProductID),
               'May')
p2.addConstrs((y[i, t] == yp[i][t - 1]*d[i][t-1] - yn[i][t-1]*(1-d[i][t-1])*Backorder_percentage[i] + quicksum(x[i][k][t - k -1] for k in Shipping_method) - Demandlist[i, t]
              for i in ProductID for t in range(3, 6)), "ending inventory")


'''
# p2.optimize()

