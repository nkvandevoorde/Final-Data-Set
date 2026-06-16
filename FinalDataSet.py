from openpyxl import workbook
from openpyxl.styles import PatternFill
import openpyxl
from openpyxl import load_workbook
import unicodedata
import pandas as pd
from rapidfuzz import fuzz
import re

cnes_wb = load_workbook("Final_CNES_data_Cleaned.xlsx").active
rel_wb = load_workbook("Clean_Relatorio_Book.xlsx").active

# #Print column names in both datasets
# print("Columns in CNES dataset:", [cell.value for cell in cnes_wb[1] if cell.value is not None])
# print("Columns in Relatório dataset:", [cell.value for cell in rel_wb[1] if cell.value is not None])

##Relatório Data Processing

# #Concatenate the Relatório Number, Addresss, and Zip Code columns into one "Site Address" Column
# site_address_col = rel_wb.max_column+1
# rel_wb.cell(row=1, column=site_address_col).value = "Site Address"

# for excel_row, row in enumerate(rel_wb.iter_rows(min_row=2, values_only=True), start=2):
#     rel_number = row[8]  # Relatório number in column I
#     address = row[9]  # Address in column J
#     zip_code = row[10]  # Zip Code in column K
#     site_address = f"{rel_number} - {address} - {zip_code}"
#     rel_wb.cell(row=excel_row, column=site_address_col).value = site_address

# #Compare CEP and Zip Code columns
# highlight = PatternFill(fill_type="solid", fgColor="d17669")
# mismatch_count = 0

# for row in range(2, rel_wb.max_row + 1):
#     cep = str(rel_wb.cell(row=row, column=4).value).strip()  # CEP in column D
#     zip_code = str(rel_wb.cell(row=row, column=11).value).strip()  # Zip Code in column K
#     if cep != zip_code:
#         rel_wb.cell(row=row, column=4).fill = highlight
#         rel_wb.cell(row=row, column=11).fill = highlight
#         mismatch_count += 1

# print(f"Number of mismatches: {mismatch_count}")
# #519/2616 = 19.839% Zip Code Mismatches
# rel_wb.parent.save("Clean_Relatorio_Book.xlsx")

# #Comparing addresses in Relatório dataset
# def clean_address(address):
#     if address is None:
#         return ""
    
#     address = str(address).upper()

#     # Remove accents
#     address = ''.join(c for c in unicodedata.normalize('NFD', address) if unicodedata.category(c) != 'Mn')

#     #Remove Punctuation
#     address = re.sub(r'[^A-Z0-9\s]', ' ', address)

#     #Remove extra spaces
#     address = re.sub(r'\s+', ' ', address).strip()

#     return address

# orange = PatternFill(fill_type="solid", fgColor="c28651")
# mismatch_amt = 0

# for row in range(2, rel_wb.max_row + 1):
#     sheets_address = clean_address(rel_wb.cell(row=row, column=5).value)  # Address in column E
#     site_address = clean_address(rel_wb.cell(row=row, column=12).value)  # Site Address in column L

#     similarity = fuzz.token_set_ratio(sheets_address, site_address)

#     #Threshold
#     if similarity < 60:  # Adjust the threshold as needed
#         rel_wb.cell(row=row, column=5).fill = orange
#         rel_wb.cell(row=row, column=12).fill = orange
#         mismatch_amt += 1

# print(f"Number of address mismatches: {mismatch_amt}")
# #830/2616 = 31.72% (roughly)Address Mismatches
# rel_wb.parent.save("Clean_Relatorio_Book.xlsx")

#Create new sheet to combine relevant data
new_sheet = rel_wb.parent.create_sheet("Final Data")
#add columns
new_sheet.append(["CNES", "State", "City", "ZIP Code (sheet)", "ZIP Code (site)", "Address(sheet)", "Address(site)", "Latitude (Relatorio)", 
                  "Longitude (Relatorio)", "Latitude (CNES)", "Longitude (CNES)", "REGIC Label"])

#Check for matching CNES numbers between the two datasets
#Extract CNES and Relatório numbers
cnes_numbers = [str(cell.value).strip() for cell in cnes_wb['C'] if cell.value is not None] #CNES in column C
rel_numbers = [str(cell.value).strip() for cell in rel_wb['B'] if cell.value is not None] #Relatório in column B
#Find matches and outliers
matching_numbers = set(cnes_numbers) & set(rel_numbers)
cnes_outliers = set(cnes_numbers) - set(rel_numbers)
rel_outliers = set(rel_numbers) - set(cnes_numbers)
print("Number of matching CNES numbers:", len(matching_numbers))
print("Number of CNES outliers:", len(cnes_outliers))
print("Number of Relatório outliers:", len(rel_outliers))

#Create Library for CNES sheet
cnes_library = {}
for excel_row, row in enumerate(cnes_wb.iter_rows(min_row=2, values_only=True), start=2):
    cnes_library[excel_row] = {
        "State": row[0],  # State in column A
        "City": row[1],   # City in column B
        "CNES": row[2],   # CNES in column C
        "Latitude": row[3], # Latitude in column D
        "Longitude": row[4], # Longitude in column E
        "REGIC Label": row[5], # REGIC Label in column F
        "Address": row[6], # Address in column G
        }
print(list(cnes_library.items())[:5])

#Create Library for Relatório sheet
rel_library = {}
for excel_row, row in enumerate(rel_wb.iter_rows(min_row=2, values_only=True), start=2):
        rel_library[excel_row] = {
            "State": row[0],  # State in column A
            "CNES": row[1],   # CNES in column B
            "City": row[2],   # City in column C
            "ZIP Code (sheet)": row[3], # ZIP Code in column D
            "Address (sheet)": row[4], # Address in column E
            "Latitude": row[6], # Latitude in column G
            "Longitude": row[7], # Longitude in column H
            "ZIP Code (site)": row[10], # ZIP Code in column K
            "Address (site)": row[11], # Address in column L
        }

print(list(rel_library.items())[:5])

#rel_wb.parent.save("Clean_Relatorio_Book.xlsx")

