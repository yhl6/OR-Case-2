import time
import pandas as pd
from pathlib import Path
import numpy as np

'''
讀取資料
'''
DATA_PATH = Path(__file__).resolve().parent / 'data'
file_name = DATA_PATH / 'OR109-2_case02_data_s2.xlsx'
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
Conflict_list = [*zip(Conflict['Product 1'], Conflict['Product 2'])]
Fixed_cost = [100, 80, 50]

N, M = Demand.shape  # 產品數、月份數
R = Vendor_cost.shape[0]  # 供應商的數量
J = Conflict.shape[0]  # 衝突的數量(共有幾種衝突)

ProductID = range(1, N + 1)
MonthID = range(1, M + 1)
Shipping_method = range(1, 4)
VendorID = range(1, R + 1)
ConflictID = range(1, J + 1)

row_index = pd.Index([*range(1, N + 1)], name='Product')
column_names = pd.MultiIndex.from_product([MonthID, ['E', 'A', 'O']], names=['Month', 'Method'])
sol_df = pd.DataFrame(0, index=row_index, columns=column_names)

sc_arr = pd.Series(0, index=ProductID)
Letter = {1: 'E', 2:'A', 3:'O'}
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
                elif (Lb[i] >= demand_unsati + 100):  # 如果沒有低於下限很多就訂到下限(如果低於下限很多就乾脆不訂)
                    sol_df.loc[i, (t, 'E')] = Lb[i]
                endinv[t + 1] = 0
            else:  # 不用order
                endinv[t + 1] = -demand_unsati

            # if (Lb[i] <= Demand.loc[i, t + 1]):  # 如果高於下限就直接訂
            #     sol_df.loc[i, (t, 'E')] = Demand.loc[i, t + 1]
            # elif (Lb[i] >= Demand.loc[i, t + 1] + 100):  # 如果沒有低於下限很多就訂到下限(如果低於下限很多就乾脆不訂)
            #     sol_df.loc[i, (t, 'E')] = Lb[i]

for j in ConflictID:
    conflict_1 = Conflict.loc[j, 'Product 1']
    conflict_2 = Conflict.loc[j, 'Product 2']

    for t in MonthID[:-1]:
        if (sum(sol_df.loc[conflict_1, t]) * sum(sol_df.loc[conflict_2, t]) != 0):
            for k in Shipping_method:
                if sc_arr[conflict_1] >= sc_arr[conflict_2]:  # 1的後加到前，2的前加到後
                    sol_df.loc[conflict_1, (t, Letter[k])] += sol_df.loc[conflict_1, (t + 1, Letter[k])]
                    sol_df.loc[conflict_1, (t + 1, Letter[k])] = 0
                    sol_df.loc[conflict_2, (t + 1, Letter[k])] += sol_df.loc[conflict_2, (t, Letter[k])]
                    sol_df.loc[conflict_2, (t, Letter[k])] = 0
                else:
                    sol_df.loc[conflict_2, (t, Letter[k])] += sol_df.loc[conflict_2, (t + 1, Letter[k])]
                    sol_df.loc[conflict_2, (t + 1, Letter[k])] = 0
                    sol_df.loc[conflict_1, (t + 1, Letter[k])] += sol_df.loc[conflict_1, (t, Letter[k])]
                    sol_df.loc[conflict_1, (t, Letter[k])] = 0
# sol_df['sum'] = sol_df.apply(lambda  x:x.sum(), axis=1)
# print(sol_df)
sol_df.to_excel('Heu Result.xlsx')
