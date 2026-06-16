from openpyxl import workbook
import openpyxl
from openpyxl import load_workbook
import unicodedata
import pandas as pd

cnes_wb = load_workbook("Final_CNES_data.xlsx").active
rel_wb = load_workbook("Final_Relatorio_data.xlsx").active

#Step 1: get rid of unnecessary columns
#Delete columns A-C, G-H, K-L
cnes_wb.delete_cols(1, 3)  # Delete columns A to C
cnes_wb.delete_cols(7, 8)  # Delete columns G to H
cnes_wb.delete_cols(11, 12)  #Delete columns K to L

#Delete columns A-E, G-L, N, P-Q, S, V, W-AL
rel_wb.delete_cols(1, 5)  # Delete columns A to E
rel_wb.delete_cols(7, 12) #Delete columns G to L
rel_wb.delete_cols(14) #Delete columns N
rel_wb.delete_cols(16, 17) #Delete columns P and Q
rel_wb.delete_cols(19) #Delete columns S
rel_wb.delete_cols(22) #Delete columns V
rel_wb.delete_cols(23, 38) #Delete columns W to AL

#Check for matching CNES numbers between the two datasets
cnes_numbers = [cell.value for cell in cnes_wb['F'] if cell.value is not None]
rel_numbers = [cell.value for cell in rel_wb['M'] if cell.value is not None]
matching_numbers = set(cnes_numbers) & set(rel_numbers)
print("Matching CNES numbers:", len(matching_numbers))
