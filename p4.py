from gurobipy import *
import pandas as pd
import numpy as np
import time


# pd.set_option('display.max_columns', None, 'display.max_rows', None)
optimal_sol = [1292791374.341664, 172131608.18494081, 578510781.5986995, 919009002.4521877, 1153364296.5329633]
'''
讀取資料
'''

def heuristic_algorithm(num):
    order = []
    start = time.time()
    file_path = 'data/'+ f'OR109-2_case02_data_s{num}.xlsx'
    xls = pd.ExcelFile(file_path)
    Demand = pd.read_excel(xls, sheet_name='Demand', index_col=0, dtype=np.int64)

    # 期初存貨
    Initial_inventory = pd.read_excel(xls, sheet_name='Initial inventory', index_col=0, dtype=np.int64)
    Initial_inv = Initial_inventory['Initial inventory']

    # 變動運輸成本
    Shipping_cost = pd.read_excel(xls, sheet_name='Shipping cost', index_col=0, dtype=np.int64)
    Express_delivery = Shipping_cost['Express delivery']
    Air_freight = Shipping_cost['Air freight']

    # 在途存貨
    In_transit = pd.read_excel(xls, sheet_name='In-transit', index_col=0, dtype=np.int64)

    # 產品體積
    Size = pd.read_excel(xls, sheet_name='Size', index_col=0, dtype=np.float64)
    Cubic_meter = Size['Size']

    # 售價、購買成本、持有成本
    Price_and_cost = pd.read_excel(xls, sheet_name='Price and cost', index_col=0, dtype=np.float64)
    Purchasing_cost = Price_and_cost['Purchasing cost']
    Holding_cost = Price_and_cost['Holding cost']

    # 銷售順延成本、銷售損失成本、銷售順延比率
    Shortage = pd.read_excel(xls, sheet_name='Shortage', index_col=0, dtype=np.float64)
    Backorder = Shortage['Backorder']  # 0.05 p_i
    Lost_sales = Shortage['Lost sales']  # 機會成本
    Bo_percent = Shortage['Backorder percentage']  # 願意繼續等的機率

    # 其他資料

    Vendor_cost = pd.read_excel(xls, sheet_name='Vendor cost', index_col=0, dtype=np.int64)

    Bounds = pd.read_excel(xls, sheet_name='Bounds', index_col=0, dtype=np.int64)
    Lb = Bounds['Minimum order quantity (if an order is placed)']
    Conflict = pd.read_excel(xls, sheet_name='Conflict', index_col=0)

    N, M = Demand.shape  # 產品數、月份數
    R = len(Vendor_cost.index)  # 供應商的數量
    J = len(Conflict.index)  # 衝突的數量(共有幾種衝突)
    ProductID = range(1, N + 1)
    MonthID = range(1, M + 1)
    Shipping_method = range(1, 4)
    ConflictID = range(1, J + 1)

    # TD[i] = Total demand for product i
    row_index = pd.Index([*range(1, N + 1)], name='Product')
    column_names = pd.MultiIndex.from_product([MonthID, ['E', 'A', 'O']], names=['Month', 'Method'])
    sol_df = pd.DataFrame(0, index=row_index, columns=column_names)

    sc_arr = pd.Series(0, index=ProductID)
    for i in ProductID:
        endinv = pd.Series(0, index=MonthID)
        # c1 = Express_delivery[i] + Holding_cost[i] * 2 + 100
        c2 = Air_freight[i] + Holding_cost[i] + 80
        c3 = (Cubic_meter[i] / 0.5) * 1500 + 50
        sc = Backorder[i] * Bo_percent[i] + Lost_sales[i] * (1 - Bo_percent[i])
        sc_arr[i] = sc
    
        if (c2 <= c3):
            cost, method = c2 + Purchasing_cost[i], 2
        else:
            cost, method = c3 + Purchasing_cost[i], 3
        # print(i,cost - sc)
        # 第一期怎麼定(如果第二期會不滿足則有用陸運)(假設前幾期一定要滿足所以才會訂)
        endinv[1] = Initial_inv[i] + In_transit.loc[i, 1] - Demand.loc[i, 1]
        endinv[2] = max(endinv[1], 0) + min(0, endinv[1]) * Bo_percent[i] + In_transit.loc[i, 2] - Demand.loc[i, 2]
        if endinv[2] < 0:
            sol_df.loc[i, (1, 'E')] = max(-endinv[2], Lb[i])
            endinv[2] = 0

        endinv[3] = max(endinv[2], 0) + min(0, endinv[2]) * Bo_percent[i] + In_transit.loc[i, 3] - Demand.loc[i, 3]
        if endinv[3] < 0:
            sol_df.loc[i, (1, 'A')] = max(-endinv[3], Lb[i])
            endinv[3] = max(0, Lb[i] - max(-endinv[3], Lb[i]))

    for i in ProductID:
        sc = sc_arr[i]
        for t in MonthID:
            if method == 3:  # t期訂 t+3期用
                if Bo_percent[i] != 1:
                    compounded_sc = sc * ((1 - (Bo_percent[i] ** (M - t - 2))) / (1 - Bo_percent[i]))
                else:
                    compounded_sc = sc * (M - t - 2)
                # print(i, t, compounded_sc)
                if (compounded_sc <= cost) | (t >= 24):
                    break
                demand_unsati = Demand.loc[i, t + 3] - endinv[t + 2]
                if demand_unsati > 0:  # 需要order
                    if Lb[i] <= demand_unsati:  # 如果高於下限就直接訂
                        sol_df.loc[i, (t, 'O')] = demand_unsati
                        endinv[t + 3] = 0
                    elif Lb[i] >= demand_unsati + 200:  # 如果沒有低於下限很多就訂到下限(如果低於下限很多就乾脆不訂)
                        sol_df.loc[i, (t, 'O')] = Lb[i]
                        endinv[t + 3] = Lb[i] - demand_unsati
                else:  # 不用order
                    endinv[t + 3] = -demand_unsati
            # elif method == 2:
            else:
                if t < 2:
                    continue
                if Bo_percent[i] != 1:
                    compounded_sc = sc * ((1 - (Bo_percent[i] ** (M - t - 1))) / (1 - Bo_percent[i]))
                else:
                    compounded_sc = sc * (M - t - 1)
                # print(i, t, compounded_sc)
                if (compounded_sc <= cost) | (t >= 25):
                    break
                demand_unsati = Demand.loc[i, t + 2] - endinv[t + 1]
                if demand_unsati > 0:  # 需要order
                    if Lb[i] <= demand_unsati:  # 如果高於下限就直接訂
                        sol_df.loc[i, (t, 'A')] = demand_unsati
                        endinv[t + 2] = 0
                    elif Lb[i] >= demand_unsati + 200:  # 如果沒有低於下限很多就訂到下限(如果低於下限很多就乾脆不訂)
                        sol_df.loc[i, (t, 'A')] = Lb[i]
                        endinv[t + 2] = Lb[i] - demand_unsati
                else:  # 不用order
                    endinv[t + 2] = -demand_unsati
            # else:
            #     if t < 3:
            #         continue
            #
            #     if Bo_percent[i] != 1:
            #         compounded_sc = sc * ((1 - (Bo_percent[i] ** (M - t))) / (1 - Bo_percent[i]))
            #     else:
            #         compounded_sc = sc * (M - t)
            #
            #     if (compounded_sc <= cost) | (t >= 25):
            #         break
            #
            #     demand_unsati = Demand.loc[i, t + 1] - endinv[t]
            #     if demand_unsati > 0:  # 需要order
            #         if Lb[i] <= demand_unsati:  # 如果高於下限就直接訂
            #             sol_df.loc[i, (t, 'E')] = demand_unsati
            #             endinv[t + 1] = 0
            #         elif Lb[i] >= demand_unsati + 200:  # 如果沒有低於下限很多就訂到下限(如果低於下限很多就乾脆不訂)
            #             sol_df.loc[i, (t, 'E')] = Lb[i]
            #             endinv[t + 1] = Lb[i] - demand_unsati
            #     else:  # 不用order
            #         endinv[t + 1] = -demand_unsati

    # 處理衝突
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
                else:
                    if sc_arr[conflict_1] >= sc_arr[conflict_2]:
                        product_to_stay, product_to_delay = conflict_1, conflict_2
                    else:
                        product_to_stay, product_to_delay = conflict_2, conflict_1

                for val in ['E', 'A', 'O']:
                    sol_df.loc[product_to_stay, (t, val)] += sol_df.loc[product_to_stay, (t + 1, val)]
                    sol_df.loc[product_to_stay, (t + 1, val)] = 0
                    sol_df.loc[product_to_delay, (t + 1, val)] += sol_df.loc[product_to_delay, (t, val)]
                    sol_df.loc[product_to_delay, (t, val)] = 0
    end = time.time()
    # for i in ProductID:
    #     if sol_df.loc[i].sum() == 0:
    #         coooost = min(Air_freight[i] + Holding_cost[i] + 80, (Cubic_meter[i] / 0.5) * 1500 + 50)
    #         sccccc =  Backorder[i] * Bo_percent[i] + Lost_sales[i] * (1 - Bo_percent[i])
    #         print(i, coooost - sccccc)

    print(f's{num}\'s time: {end - start}')
    print('-------------')

heuristic_algorithm(1)
heuristic_algorithm(2)
heuristic_algorithm(3)
heuristic_algorithm(4)
heuristic_algorithm(5)