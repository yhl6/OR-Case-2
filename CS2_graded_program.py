'''
You do not need to change the code in this file.
You only need to ensure we could run your algorithm here.
'''
import time
import pandas as pd
import numpy as np
import os
import gurobipy as gp
from gurobipy import GRB
from algorithm_module import heuristic_algorithm

if __name__ == '__main__':
    
    # read all instances (xlsx file) under data folder
    all_data_list = os.listdir('data')
    
    # evaluate all instances 
    result_df = pd.DataFrame(columns=['Data name', 'Time', 'Objective value', 'Feasibility'])
    
    for file_name in all_data_list:
        
        start_time = time.time()        
        try:
            '''
            1. We will import your algorithm here and give you file_path (e.g.,'data/OR109-2_case02_data_s1.xlsx') as function argument.
               
            2. You need to return the order planning in list in order.
                    order[i][j][t] is amount of product i ordered in the beginning of month t with shipping method j.
                    i = 1, ..., |number of product|, j = 1, 2, 3, and t = 1, ..., 26.
               Note that the indice of list need to be in order, that is , return order[i][j][t] rather than order[t][j][i]
            '''
            file_path = 'data/'+file_name
            order = heuristic_algorithm(file_path)
            
            '''
            We would calculate objective value and check feasibility here.
            '''
            # feasibility, obj_value = check_feasibility_and_calculate_objective_value(file_path, order)
            
        except:
            print("the algorithm has errors")
            # obj_value = np.nan
            # feasibility = False

        # end_time = time.time()
        # spend_time = end_time - start_time
        
        # result_df = result_df.append({'Data name': file_name, 
        #                               'Time': spend_time,
        #                               'Objective value': obj_value, 
        #                               'Feasibility': feasibility},
        #                             ignore_index=True)

# output result
# result_df.to_csv('result.csv', index=False)