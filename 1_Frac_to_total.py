import xlsxwriter
from win32com.client.gencache import EnsureDispatch
import os
import re
from pprint import pprint
import pandas as pd
from functools import reduce
import itertools
import sqlite3
import win32com
print(win32com.__gen_path__)

# Excel connection  
xl = EnsureDispatch('Excel.Application')
wb_total = xl.Workbooks.Open(f"{os.getcwd()}\\СВОДНАЯ СВЕРКА.xlsx")
ws_total = wb_total.Worksheets(2)
wb_frac = xl.Workbooks.Open(f"{os.getcwd()}\\Сверка ГРП на 31.08.2022.xlsx")
ws_frac = wb_frac.Worksheets(1)

# Pandas
pd.set_option('display.max_rows', None)

# db connections
db = sqlite3.connect('total_match.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

def main():
    # Get plates from Total match xls and turn them into 123abc type as L_tm_plate_ind (length 307)
    def platesTotal(row, col):
        L_total_plates = []
        while row != 309:
            L_total_plates.append(ws_total.Cells(row, col).Value)
            row += 1
        return L_total_plates
            
    L_total_plates = platesTotal(2, 5)
    
    # Get all cols from Frac rep (length 485)
    def dataFrac(row, col):
        L_data = []
        while row != 490:
            L_data.append(ws_frac.Cells(row, col).Value)
            row += 1
        return L_data
            
    L_frac_group = dataFrac(5, 2)
    L_frac_unit = dataFrac(5, 3)
    L_frac_plates = dataFrac(5, 4)
    L_frac_mols = dataFrac(5, 5)
    L_frac_drivers = dataFrac(5, 6)
    L_frac_discrepancies = dataFrac(5, 7)
    L_frac_notes = dataFrac(5, 8)
    
    
    
    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions = [186, 86, 797, '02', '07', 82, 78, 54, 77, 126, 188, 88, 174, 74, 158, 196, 156, 76, 1]
        
        for i in L_regions:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) > 7 else x for x in plates]
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() if x != None else None for x in plates ]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() if x != None else None for x in plates]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() if x != None else None for x in plates]

        return plates
    
    L_frac_plates = [str(x) for x in L_frac_plates]
    L_total_plates_ind = transform_plates(L_total_plates) 
    L_total_frac_ind = transform_plates(L_frac_plates)

    # Matching multiple columns via dict
    def dict_matcher(x):
        L = []
        D = dict(zip(L_total_frac_ind, x))
        for i in L_total_plates_ind:
            L.append(D.get(i))
        return L
    
    L_group_matched = dict_matcher(L_frac_group)
    L_unit_matched = dict_matcher(L_frac_unit)    
    L_plates_matched = dict_matcher(L_frac_plates)    
    L_mols_matched = dict_matcher(L_frac_mols)    
    L_drivers_matched = dict_matcher(L_frac_drivers)    
    L_discrepancies_matched = dict_matcher(L_frac_discrepancies)    
    L_notes_matched = dict_matcher(L_frac_notes)    
    
    # def post_to_xls(L, col):
    #     row = 2
    #     for i in L:
    #         ws_total.Cells(row, col).Value = i    
    #         row += 1
    
    # L_category = [L_group_matched, L_unit_matched, L_plates_matched, L_mols_matched, L_drivers_matched, L_discrepancies_matched, L_notes_matched]
    # L_col = [8, 9, 10, 12, 13, 14, 15]

    # for i, j in zip(L_category, L_col):
    #     post_it = post_to_xls(i, j)
    
    

    # data = pd.DataFrame(zip(L_group_matched, L_unit_matched, L_plates_matched, L_mols_matched, L_drivers_matched, L_discrepancies_matched, L_notes_matched))

    
    # Match if frac items are not in Omnicomm (see accountance algo)
    def dict_unmatcher(x):
        L = []
        D = dict(zip(L_total_frac_ind, x))
        for i in L_total_frac_ind:
            if i not in L_total_plates_ind:
                L.append(D.get(i))
        return L
    
    L_group_unmatched = dict_unmatcher(L_frac_group)
    L_unit_unmatched = dict_unmatcher(L_frac_unit)    
    L_plates_unmatched = dict_unmatcher(L_frac_plates)    
    L_mols_unmatched = dict_unmatcher(L_frac_mols)    
    L_drivers_unmatched = dict_unmatcher(L_frac_drivers)    
    L_discrepancies_unmatched = dict_unmatcher(L_frac_discrepancies)    
    L_notes_unmatched = dict_unmatcher(L_frac_notes)

    def post_to_xls2(L, col):
        row = 309
        for i in L:
            ws_total.Cells(row, col).Value = i    
            row += 1
    
    L_category = [L_group_unmatched, L_unit_unmatched, L_plates_unmatched, L_mols_unmatched, L_drivers_unmatched, L_discrepancies_unmatched, L_notes_unmatched]
    L_col = [8, 9, 10, 12, 13, 14, 15]

    for i, j in zip(L_category, L_col):
        post_it = post_to_xls2(i, j)
    
    wb_total.Close(True)
    wb_frac.Close(True)
    xl.Quit()
if __name__ == '__main__':
    main()