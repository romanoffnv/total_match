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
    
    wb_total.Close(True)
    wb_frac.Close(True)
    xl.Quit()
    
    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions = [186, 86, 797, '02', '07', 82, 78, 54, 77, 126, 188, 88, 174, 74, 158, 196, 156, 76, 1]
        
        for i in L_regions:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) > 7 else x for x in plates]
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]

        return plates
    
    L_frac_plates = [str(x) for x in L_frac_plates]
    L_total_plates_ind = transform_plates(L_total_plates) 
    L_total_frac_ind = transform_plates(L_frac_plates)
    
    # make df of all cols and plate index
    data = pd.DataFrame(zip(L_frac_group, L_frac_unit, L_frac_plates, L_total_frac_ind, L_frac_mols, L_frac_drivers, L_frac_discrepancies, L_frac_notes))
    print(data)
    print(data.describe())
    
    # filter df by L_tm_plate_ind (to match for omnicomm)
    # Match if frac items are not in Omnicomm (see accountance algo)

if __name__ == '__main__':
    main()