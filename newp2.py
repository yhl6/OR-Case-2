from gurobipy import *
import pandas as pd
from pathlib import Path
import numpy as np
import time
import locale

pd.set_option('display.max_columns', None, 'display.max_rows', None)
# locale.setlocale( locale.LC_ALL, '' )
'''
讀取資料
'''
for num in [2]:
    start = time.time()
    DATA_PATH = Path(__file__).resolve().parent / 'data'
    file_name = DATA_PATH / f'OR109-2_case02_data_s{num}.xlsx'
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
    ProductID_grouped = Vendor_product.groupby('Vendor')

    Vendor_cost = pd.read_excel(df, sheet_name='Vendor cost', index_col=0)
    Ordering_cost = Vendor_cost['Ordering cost']

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

    # TD[i] = Total demand for product i
    Total_demand = sum(Demand.apply(lambda x: x.sum(), axis=1))

    # 模型的部分
    p2 = Model("p2")
    p2.Params.LogToConsole = 0
    # Non-binary Variables
    x = p2.addVars(ProductID, Shipping_method, range(0, M + 1), lb=0, vtype=GRB.INTEGER, name='x')  # x[i, k, t]
    v = p2.addVars(ProductID, range(0, M + 1), lb=0, vtype= GRB.CONTINUOUS, name= 'v')  # 期末存貨 v[i, t]
    s = p2.addVars(ProductID, range(0, M + 1), lb=0, vtype= GRB.CONTINUOUS, name= 's')  # Shortage s[i, t]
    g = p2.addVars(MonthID, lb=0, vtype=GRB.INTEGER, name='g')  # 高斯符號 g[t]
    D = p2.addVars(ProductID, MonthID, vtype=GRB.CONTINUOUS, name='D')  # Demand of product i in week t as considering backorder

    # Binary Variables
    z = p2.addVars(Shipping_method, MonthID, lb=0, vtype=GRB.BINARY, name='z')  # z[k, t] = 1 if method k is used in month t
    d = p2.addVars(ProductID, MonthID, lb=0, vtype=GRB.BINARY, name='d')  # d[i, t] = 1 if v[i, t] = 0 or = 0 otherwise.
    a = p2.addVars(ProductID, MonthID, lb=0, vtype=GRB.BINARY, name='a')  # a[i, t] = 1 if product i is ordered in month t or = 0 otherwise.
    o = p2.addVars(VendorID, MonthID, lb=0, vtype=GRB.BINARY, name='o')  # o[r, t] = 1 if we order any product from vendor r

    # 更新變數
    p2.update()

    # 目標式

    p2.setObjective(
        quicksum(Express_delivery[i] * quicksum(x.select(i, 1, '*')) for i in ProductID)  # 陸運變動成本
        + quicksum(Air_freight[i] * quicksum(x.select(i, 2, '*')) for i in ProductID)  # 空運變動成本
        + quicksum(Fixed_cost[k - 1] * quicksum(z.select(k, '*')) for k in Shipping_method)  # 固定成本
        + quicksum(Purchasing_cost[i] * quicksum(x.select(i, '*', '*')) for i in ProductID)  # 購買成本
        + quicksum(Holding_cost[i] * quicksum(v[i, t] for t in MonthID) for i in ProductID)  # 倉儲成本
        + 1500 * quicksum(g.select())  # 貨櫃成本
        + quicksum(quicksum(Lost_sales[i] * s[i, t] * (1 - Backorder_percentage[i]) for i in ProductID) for t in MonthID)  # Lost sales cost
        + quicksum(quicksum(Backorder[i] * Backorder_percentage[i] * s[i, t] for i in ProductID) for t in MonthID)  # Backorder cost
        + quicksum(Ordering_cost[r] * quicksum(o[r, t] for t in MonthID) for r in VendorID), GRB.MINIMIZE)  # Vendor cost


    # 限制式

    p2.addConstrs((v[i, 0] == 0) for i in ProductID)
    p2.addConstrs((s[i, 0] == 0) for i in ProductID)

    for i in ProductID:
        for k in Shipping_method:
            p2.addConstr(x[i, k, 0] == 0)

    # y[i, t] Constraint
    for i in ProductID:
        p2.addConstr((v[i, 1] == v[i, 0] + s[i, 1] + In_transit.loc[i, 1] - D[i, 1] + Initial_inv[i]), 'Week 1')
        p2.addConstr((v[i, 2] == v[i, 1] + s[i, 2] + In_transit.loc[i, 2] - D[i, 2] + x[i, 1, 1] ), 'Week 2')
        p2.addConstr((v[i, 3] == v[i, 2] + s[i, 3] + In_transit.loc[i, 3] - D[i, 3] + x[i, 1, 2] + x[i, 2, 1] ), 'Week 3')

    for i in ProductID:
        for t in MonthID[3:]:
            p2.addConstr((v[i, t] == v[i, t - 1] + s[i, t] + x[i, 1, t - 1] + x[i, 2, t - 2] + x[i, 3, t - 3] - D[i, t]), 'Week others')

    # Z Constraint
    for t in MonthID:
        for k in Shipping_method:
            p2.addConstr(quicksum(x[i, k, t] for i in ProductID) <= Total_demand * z[k, t], 'Z constraint')

    # Y的binary constraint
    for i in ProductID:
        for t in MonthID:
            p2.addConstr(Total_demand * d[i, t] >= s[i, t])
            p2.addConstr(Total_demand * (1 - d[i, t]) >= v[i, t])
            p2.addConstr(D[i, t] >= Demand.loc[i, t] + Backorder_percentage[i] * s[i, t - 1])

    # ceiling constraints
    for t in MonthID:
        p2.addConstr(g[t] >= quicksum(Cubic_meter[i] * x[i, 3, t] for i in ProductID) / 0.5)

    # conflict constraint
    for i, j in Conflict_tl:
        for t in MonthID:
            p2.addConstr((a[i, t] + a[j, t] <= 1))

    # o[r, t] constraints
    for r in VendorID:
        for t in MonthID:
            for i in [ProductID_grouped.groups[r]]:
                p2.addConstr((quicksum(x.select(i, '*', t))) <= Total_demand * o[r, t])

    # Lower Bound constraint
    p2.addConstrs((quicksum(x.select(i, '*', t)) >= Lower_bound[i] * a[i, t]) for i in ProductID for t in MonthID)
    p2.addConstrs((quicksum(x.select(i, '*', t)) <= Total_demand * a[i, t]) for i in ProductID for t in MonthID)

    p2.optimize()

    # print(f's{num}\'s objVal: {locale.currency(p2.objVal, grouping=True)}')
    print(p2.objVal)

    row_index = pd.Index([*range(1, N + 1)], name='Product')
    column_names = pd.MultiIndex.from_product([MonthID, ['E', 'A', 'O']], names=['Month', 'Method'])
    sol_df = pd.DataFrame(index=row_index, columns=column_names)
    print(sol_df)


    for i in ProductID:
        for t in MonthID:
            for k in Shipping_method:
                if k == 1:
                    sol_df.loc[i, (t, 'E')] = x[i, k, t].x
                elif k == 2:
                    sol_df.loc[i, (t, 'A')] = x[i, k, t].x
                elif k == 3:
                    sol_df.loc[i, (t, 'O')] = x[i, k, t].x
    # print(sol_df)
    # sol_df.to_excel(f'Result_{num}.xlsx', sheet_name='Ordering Plan')
    # for k in Shipping_method:
    #     for t in MonthID:
    #         if z[k, t].x != 0:
    #             print(f'z[{k}, {t}]: {z[k, t].x}')

    # for r in VendorID:
    #     print(r, Ordering_cost[r] * (26 - quicksum((val.x for val in o.select(r, '*')))))
