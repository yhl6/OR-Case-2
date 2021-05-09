from gurobipy import *
import pandas as pd
import time
import math
import numpy as np

def heuristic_algorithm(file_path,n):

	start = time.time()
	local_time = time.ctime(start)
	df = pd.read_excel(file_path)
	
	#需求
	Demand_info = pd.read_excel(file_path, 'Demand')
	Demandlist = []
	for i in Demand_info.index:
		Demandlist.append(list(Demand_info.loc[i]))

	for i in range(len(Demandlist)):
		Demandlist[i].pop(0)
	#print(Demandlist)

	# 產品、月份、運送方法編號 i k t
	ProductID = range(len(Demand_info['Product']))
	Shipping_method = range(0, 3)
	MonthID = range(0, 26)

	# 初始存貨
	Initial_Inv = pd.read_excel(file_path, 'Initial inventory')
	Initial_Inventory = Initial_Inv['Initial inventory']
	#print(Initial_Inventory)

	# 運送費用
	Shipping_Cost = pd.read_excel(file_path, 'Shipping cost')
	Express_delivery = Shipping_Cost['Express delivery']
	Air_freight = Shipping_Cost['Air freight']
	Fixed_cost = [100, 80, 50]

	# 在途存貨
	In_transit = pd.read_excel(file_path, 'In-transit')
	April_intransit = In_transit[2]
	May_intransit = In_transit[3]

	# 箱子大小(第二題)
	Size = pd.read_excel(file_path, 'Size')
	Cubic_meter = Size['Size']

	# 售價、購買成本、期末持有成本
	Price_info = pd.read_excel(file_path, 'Price and cost')
	Sales = Price_info['Sales price']
	Purchasing_cost = Price_info['Purchasing cost']
	Holding_cost = Price_info['Holding cost']

	# Shortage & Backorder
	Shortage = pd.read_excel(file_path,'Shortage')
	Backorder = Shortage['Backorder'] # 0.05 p_i
	Lost_sales = Shortage['Lost sales'] # 機會成本
	Backorder_percentage = Shortage['Backorder percentage'] # 願意繼續等的機率

	# Bounds & Vendor
	Bounds = pd.read_excel(file_path,'Bounds')
	Minimum_order = Bounds['Minimum order quantity (if an order is placed)']
	Vendor_product = pd.read_excel(file_path,'Vendor-Product')
	VendorC = pd.read_excel(file_path,'Vendor cost')
	Vendor_cost = VendorC['Ordering cost']

	# Conflict 產品1,2
	Conflict = pd.read_excel(file_path, 'Conflict')
	Conflict_product = Conflict[['Product 1','Product 2']]


	#每期若完全不訂貨，會缺多少產品 
	x = []
	for i in ProductID:
		x.append([])

	#期末存貨，假設前提是每期的貨缺少的貨，都是當月才到，也就是說一旦缺貨，之後的期末存貨就都是0
	y = []
	for i in ProductID:
		y.append([])


	#算x,y
	Left = 0
	for i in ProductID:
		Left = 0

		if Backorder_percentage[i] >= 2:
			NotList.append(i)
		for t in MonthID:   #26 period
				
				if t == 0:
					Left = Initial_Inventory[i] - Demandlist[i][t]
				elif t == 1:
					Left = Left - Demandlist[i][t] + April_intransit[i]
				elif t == 2:
					Left = Left - Demandlist[i][t] + May_intransit[i]
				else:
					Left = Left - Demandlist[i][t]

				if Left >= 0:
					x[i].append(0)
					y[i].append(Left)
				elif Left < 0:
					x[i].append(Left*(-1))
					y[i].append(0)
					Left = 0

	#如果低於lowerbound就補
	for i in ProductID:
		for t in MonthID:
			if Minimum_order[i] > x[i][t] and x[i][t] != 0:
				x[i][t] = Minimum_order[i]
				y[i][t] += Minimum_order[i] - x[i][t]
	#Backorder
	ID = []
	for i in ProductID:
			
			sum = 0
			backsum = 0
			#and Lost_sales[i]-Purchasing_cost[i] <= 0
			for t in MonthID:
				if x[i][t] > 0 :
					#print('dif',Lost_sales[i]-Backorder[i],Backorder_percentage[i])
					#print(i,t,x[i][t] * Backorder_percentage[i]**(25-t))
					sum += x[i][t] * Backorder_percentage[i]**(25-t)
					


			#print('i',i,sum)
			if Backorder_percentage[i] != 0.0:
				constant = (Lost_sales[i]*(1-Backorder_percentage[i])+(Backorder[i]*Backorder_percentage[i]))
			elif Backorder_percentage[i] == 0.0:
				constant = Lost_sales[i]*(1-Backorder_percentage[i]) * np.sum(x[i])
				#print('c',constant)
			
			print(i,sum,sum*constant, np.sum(x[i])*Purchasing_cost[i])
			if sum*constant < np.sum(x[i])*Purchasing_cost[i] and sum != 0:
				#print('id',i)
				ID.append(i)
	print(ID)
	print(len(ID))
	ID1 = [i+1 for i in ID]
	print(ID1)
	#在backorder裡
	for i in ID:
		for t in MonthID:
			x[i][t] = 0
		#print(x[i])

	Ocean_cost = []
	for i in ProductID:
		Ocean_C = (1500/0.5) * Cubic_meter[i]
		Ocean_cost.append(Ocean_C)
	
	#船運費用比空運小就紀錄
	Ship = []
	Air = []
	for i in ProductID:
		if Ocean_cost[i] < Air_freight[i]:
			Ship.append(i)
			#print(i+1,Ocean_cost[i],Air_freight[i])

	for i in ProductID:
		if i not in Ship:
			Air.append(i)

	
	#objective values   假設所有第三周以後都用空運，除非船運比空運便宜
	z = (np.sum(Express_delivery[i]*x[i][1] for i in ProductID) 
	 + np.sum(Air_freight[i]*x[i][t] for t in range(2,26) for i in Air)
	 + 1500 * (np.sum(math.ceil((x[i][t]*Cubic_meter[i])/0.5) for i in Ship for t in range(2,26)))
	 + np.sum(Purchasing_cost[i]*x[i][t] for t in MonthID for i in ProductID)
	 + np.sum(Holding_cost[i]*y[i][t] for t in MonthID for i in ProductID)
	 + np.sum(Lost_sales[i]*(1-Backorder_percentage[i])*x[i][0] for i in ProductID)
	 + Fixed_cost[0] + Fixed_cost[1]*25 + (Vendor_cost[0] + Vendor_cost[1])*26)



	OptimalList = [1282438647,164832708,542768823,868342211,1102028614]
	print('z =',format(z,','))
	print('optimal gap =',round((z-OptimalList[n])/OptimalList[n] * 100,4),'%')

	end = time.time()
	local = time.ctime(end)
	print('start: ',local_time)
	print('end: ',local)
	print('time =',round(end-start,4),'s')
	print('-----------------------------')

file1 = 'OR109-2_case02_data_s1.xlsx'
file2 = 'OR109-2_case02_data_s2.xlsx'
file3 = 'OR109-2_case02_data_s3.xlsx'
file4 = 'OR109-2_case02_data_s4.xlsx'
file5 = 'OR109-2_case02_data_s5.xlsx'

#heuristic_algorithm(file1,0)
heuristic_algorithm(file2,1)
#heuristic_algorithm(file3,2)
#heuristic_algorithm(file4,3)
#heuristic_algorithm(file5,4)


