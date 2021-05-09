from gurobipy import *
import pandas as pd
from pathlib import Path
import time
import locale

# pd.set_option('display.max_columns', None, 'display.max_rows', None)
# locale.setlocale( locale.LC_ALL, '' )
ariels = [1292803813.469, 172446584.464, 581302506.955, 923810132.627, 1153418763.777]
'''
讀取資料
'''

def heuristic_algorithm(num):
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
    Bo_percent = Shortage['Backorder percentage']  # 願意繼續等的機率

    # 其他資料
    Vendor_product = pd.read_excel(df, sheet_name='Vendor-Product', index_col=0)
    ProductID_grouped = Vendor_product.groupby('Vendor')

    Vendor_cost = pd.read_excel(df, sheet_name='Vendor cost', index_col=0)
    Ordering_cost = Vendor_cost['Ordering cost']

    Bounds = pd.read_excel(df, sheet_name='Bounds', index_col=0)
    Lb = Bounds['Minimum order quantity (if an order is placed)']
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
    row_index = pd.Index([*range(1, N + 1)], name='Product')
    column_names = pd.MultiIndex.from_product([MonthID, ['E', 'A', 'O']], names=['Month', 'Method'])
    sol_df = pd.DataFrame(0, index=row_index, columns=column_names)

    sc_arr = pd.Series(0, index=ProductID)
    Letter = {1: 'E', 2: 'A', 3: 'O'}
    for i in ProductID:
        endinv = pd.Series(0, index=MonthID)
        c1 = Express_delivery[i] + Holding_cost[i] * 2 + 100
        c2 = Air_freight[i] + Holding_cost[i] + 80
        c3 = (Cubic_meter[i] / 0.5) * 1500 + 50
        sc_arr[i] = Backorder[i] * Bo_percent[i] + Lost_sales[i] * (1 - Bo_percent[i])
        # sc_arr[i] = sc
        if (c1 <= c2) & (c1 <= c3):
            cost, method = c1 + Purchasing_cost[i], 1
        elif (c2 <= c1) & (c2 <= c3):
            cost, method = c2 + Purchasing_cost[i], 2
        else:
            cost, method = c3 + Purchasing_cost[i], 3

        # 第一期怎麼定(如果第二期會不滿足則有用陸運)(假設前幾期一定要滿足所以才會訂)
        endinv[1] = Initial_inv[i] + In_transit.loc[i, 1] - Demand.loc[i, 1]
        endinv[2] = max(endinv[1], 0) + min(0, endinv[1]) * Bo_percent[i] + In_transit.loc[i, 2] - Demand.loc[i, 2]
        if endinv[2] < 0:
            sol_df.loc[i, (1, 'E')] = max(-endinv[2], Lb[i])
            endinv[2] = 0
        endinv[3] = max(endinv[2], 0) + min(0, endinv[2]) * Bo_percent[i] + In_transit.loc[i, 3] - Demand.loc[i, 3]
        if endinv[3] < 0:
            if ((method == 3 & (c1 <= c2)) | (method == 1)):  # 陸運比較便宜
                sol_df.loc[i, (2, 'E')] = max(-endinv[3], Lb[i])
            elif ((method == 3 & (c1 > c2)) | (method == 2)):  # 空運比較便宜
                sol_df.loc[i, (1, 'A')] = max(-endinv[3], Lb[i])
            endinv[3] = 0
        # print(i)
        # print(endinv[:3])

        for t in MonthID:
            if method == 3:  # t期訂 t+3期用
                if Bo_percent[i] != 1:
                    compounded_sc = sc_arr[i] * ((1 - (Bo_percent[i] ** (M - t - 2))) / (1 - Bo_percent[i]))
                else:
                    compounded_sc = sc_arr[i] * (M - t - 2)
                if ((compounded_sc <= cost) | (t >= 24)):
                    break
                demand_unsati = Demand.loc[i, t + 3] - endinv[t + 2]
                if demand_unsati > 0:  # 需要order
                    if (Lb[i] <= demand_unsati):  # 如果高於下限就直接訂
                        sol_df.loc[i, (t, 'O')] = demand_unsati
                        endinv[t + 3] = 0
                    elif (Lb[i] >= demand_unsati + 100):  # 如果沒有低於下限很多就訂到下限(如果低於下限很多就乾脆不訂)
                        sol_df.loc[i, (t, 'O')] = Lb[i]
                        endinv[t + 3] = 0
                else:  # 不用order
                    endinv[t + 3] = -demand_unsati
            elif method == 2:
                if t < 2:
                    continue
                if Bo_percent[i] != 1:
                    compounded_sc = sc_arr[i] * ((1 - (Bo_percent[i] ** (M - t - 1))) / (1 - Bo_percent[i]))
                else:
                    compounded_sc = sc_arr[i] * (M - t - 1)

                if ((compounded_sc <= cost) | (t >= 25)):
                    break
                demand_unsati = Demand.loc[i, t + 2] - endinv[t + 1]
                if demand_unsati > 0:  # 需要order
                    if (Lb[i] <= demand_unsati):  # 如果高於下限就直接訂
                        sol_df.loc[i, (t, 'A')] = demand_unsati
                        endinv[t + 2] = 0
                    elif (Lb[i] >= demand_unsati + 100):  # 如果沒有低於下限很多就訂到下限(如果低於下限很多就乾脆不訂)
                        sol_df.loc[i, (t, 'A')] = Lb[i]
                        endinv[t + 2] = 0
                else:  # 不用order
                    endinv[t + 2] = -demand_unsati
                # if (Lb[i] <= Demand.loc[i, t + 2]):  # 如果高於下限就直接訂
                #     sol_df.loc[i, (t, 'A')] = Demand.loc[i, t + 2]
                # elif (Lb[i] >= Demand.loc[i, t + 2] + 100):  # 如果沒有低於下限很多就訂到下限(如果低於下限很多就乾脆不訂)
                #     sol_df.loc[i, (t, 'A')] = Lb[i]
            else:
                if t < 3:
                    continue

                if Bo_percent[i] != 1:
                    compounded_sc = sc_arr[i] * ((1 - (Bo_percent[i] ** (M - t))) / (1 - Bo_percent[i]))
                else:
                    compounded_sc = sc_arr[i] * (M - t)

                if ((compounded_sc <= cost) | (t >= 25)):
                    break

                demand_unsati = Demand.loc[i, t + 1] - endinv[t]
                if demand_unsati > 0:  # 需要order
                    if (Lb[i] <= demand_unsati):  # 如果高於下限就直接訂
                        sol_df.loc[i, (t, 'E')] = demand_unsati
                        endinv[t + 1] = 0
                    elif (Lb[i] >= demand_unsati + 100):  # 如果沒有低於下限很多就訂到下限(如果低於下限很多就乾脆不訂)
                        sol_df.loc[i, (t, 'E')] = Lb[i]
                        endinv[t + 1] = 0
                else:  # 不用order
                    endinv[t + 1] = -demand_unsati

                # if (Lb[i] <= Demand.loc[i, t + 1]):  # 如果高於下限就直接訂
                #     sol_df.loc[i, (t, 'E')] = Demand.loc[i, t + 1]
                # elif (Lb[i] >= Demand.loc[i, t + 1] + 100):  # 如果沒有低於下限很多就訂到下限(如果低於下限很多就乾脆不訂)
                #     sol_df.loc[i, (t, 'E')] = Lb[i]
    # sol_df.to_excel('Before.xlsx')
    for j in ConflictID:
        conflict_1 = Conflict.loc[j, 'Product 1']
        conflict_2 = Conflict.loc[j, 'Product 2']
        for t in MonthID:
            # print(f't:{t}, sum1: {sum(sol_df.loc[conflict_1, t])}, sum2: {sum(sol_df.loc[conflict_2, t])}')
            if sum(sol_df.loc[conflict_1, t]) * sum(sol_df.loc[conflict_2, t]) != 0:
                if conflict_1 in Conflict.loc[:j - 1].values:
                    product_to_stay, product_to_delay = conflict_1, conflict_2
                elif conflict_2 in Conflict.loc[:j - 1].values:
                    product_to_stay, product_to_delay = conflict_2, conflict_1
                elif sc_arr[conflict_1] >= sc_arr[conflict_2]:
                    product_to_stay, product_to_delay = conflict_1, conflict_2
                elif sc_arr[conflict_1] < sc_arr[conflict_2]:
                    product_to_stay, product_to_delay = conflict_2, conflict_1
                # print('--------------BEFORE----------')
                # print(sol_df.loc[[conflict_1, conflict_2], [t, t + 1]])
                for k in Shipping_method:
                    sol_df.loc[product_to_stay, (t, Letter[k])] += sol_df.loc[product_to_stay, (t + 1, Letter[k])]
                    sol_df.loc[product_to_stay, (t + 1, Letter[k])] = 0
                    sol_df.loc[product_to_delay, (t + 1, Letter[k])] += sol_df.loc[product_to_delay, (t, Letter[k])]
                    sol_df.loc[product_to_delay, (t, Letter[k])] = 0
                    # if sc_arr[conflict_1] >= sc_arr[conflict_2]:  # 1的後加到前，2的前加到後
                    #     sol_df.loc[conflict_1, (t, Letter[k])] += sol_df.loc[conflict_1, (t + 1, Letter[k])]
                    #     sol_df.loc[conflict_1, (t + 1, Letter[k])] = 0
                    #     sol_df.loc[conflict_2, (t + 1, Letter[k])] += sol_df.loc[conflict_2, (t, Letter[k])]
                    #     sol_df.loc[conflict_2, (t, Letter[k])] = 0
                    # else:
                    #     sol_df.loc[conflict_2, (t, Letter[k])] += sol_df.loc[conflict_2, (t + 1, Letter[k])]
                    #     sol_df.loc[conflict_2, (t + 1, Letter[k])] = 0
                    #     sol_df.loc[conflict_1, (t + 1, Letter[k])] += sol_df.loc[conflict_1, (t, Letter[k])]
                    #     sol_df.loc[conflict_1, (t, Letter[k])] = 0
    #             print('----------AFTER-----------')
    #             print(sol_df.loc[[conflict_1, conflict_2], [t, t + 1]])
    #             print('------------')
    # print('Lb check')
    # print('-------------------')
    # for i in ProductID:
    #     for k in Shipping_method:
    #         for t in MonthID:
    #             if 0 < sol_df.loc[i, (t, Letter[k])] < Lb[i]:
    #                 print(i, k, t, Lb[i], sol_df.loc[i, (t, Letter[k])])
    # print('-------------------')
    # print('Conflict check')
    # print('-------------------')
    # for j in ConflictID:
    #     conflict_1 = Conflict.loc[j, 'Product 1']
    #     conflict_2 = Conflict.loc[j, 'Product 2']
    #     # print(f'conflict_1: {conflict_1}, conflict_2: {conflict_2}')
    #     for t in MonthID:
    #         if sum(sol_df.loc[conflict_1, t]) * sum(sol_df.loc[conflict_2, t]) != 0:
    #             print(f'!!!!!!!!!!!!!! {conflict_1}, {conflict_2} in month {t} !!!!!!!!!!!!!!!')
    # print('-------------------')
    # sol_df.to_excel('After.xlsx')


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
        + quicksum(quicksum(Lost_sales[i] * s[i, t] * (1 - Bo_percent[i]) for i in ProductID) for t in MonthID)  # Lost sales cost
        + quicksum(quicksum(Backorder[i] * Bo_percent[i] * s[i, t] for i in ProductID) for t in MonthID)  # Backorder cost
        + quicksum(Ordering_cost[r] * quicksum(o[r, t] for t in MonthID) for r in VendorID), GRB.MINIMIZE)  # Vendor cost


    # 限制式

    for i in ProductID:
        for k in Shipping_method:
            for t in MonthID:
                p2.addConstr(x[i, k, t] == sol_df.loc[i, (t, Letter[k])])

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
            p2.addConstr(D[i, t] >= Demand.loc[i, t] + Bo_percent[i] * s[i, t - 1])

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
    p2.addConstrs((quicksum(x.select(i, '*', t)) >= Lb[i] * a[i, t]) for i in ProductID for t in MonthID)
    p2.addConstrs((quicksum(x.select(i, '*', t)) <= Total_demand * a[i, t]) for i in ProductID for t in MonthID)

    p2.optimize()
    end = time.time()
    # print(f's{num}\'s objVal: {locale.currency(p2.objVal, grouping=True)}')
    print(f's{num}\'s time: {end - start}')
    print(f's{num}\'s z: {p2.objVal}')
    print(f's{num}\'s gap: {100 * (p2.objVal - ariels[num - 1]) / ariels[num - 1]}%')
    print('-------------')

heuristic_algorithm(1)
heuristic_algorithm(2)
heuristic_algorithm(3)
heuristic_algorithm(4)
heuristic_algorithm(5)