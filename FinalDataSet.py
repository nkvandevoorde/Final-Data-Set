from openpyxl import workbook
from openpyxl.styles import PatternFill
import openpyxl
from openpyxl import load_workbook
import unicodedata
import pandas as pd
from rapidfuzz import fuzz
import re

cnes_wb = load_workbook("Final_CNES_data_Cleaned.xlsx").active
rel_wb = load_workbook("/Users/noravandevoorde/Downloads/SPHERE/Clean_Relatorio_Book.xlsx").active

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
new_wb = rel_wb.parent.create_sheet("Final Data")
#add columns
new_wb.append(["CNES", "State", "City", "ZIP Code (sheet)", "ZIP Code (site)", "Address(sheet)", "Address(site)", "Latitude (Relatorio)", 
                  "Longitude (Relatorio)", "Latitude (CNES)", "Longitude (CNES)", "REGIC Label"])

#Check for matching CNES numbers between the two datasets
#Extract CNES and Relatório numbers
cnes_numbers = [str(cell.value).strip() for cell in cnes_wb['C'][1:] if cell.value is not None] #CNES in column C
rel_numbers = [str(cell.value).strip() for cell in rel_wb['B'][1:] if cell.value is not None] #Relatório in column B
#Find matches and outliers
matching_numbers = set(cnes_numbers) & set(rel_numbers)
cnes_outliers = set(cnes_numbers) - set(rel_numbers)
rel_outliers = set(rel_numbers) - set(cnes_numbers)
#print("Relatório outliers:", rel_outliers)
print("Number of matching CNES numbers:", len(matching_numbers))
print("Number of CNES outliers:", len(cnes_outliers))
print("Number of Relatório outliers:", len(rel_outliers))

#Create Library for CNES sheet
cnes_library = {}
for excel_row, row in enumerate(cnes_wb.iter_rows(min_row=2, values_only=True), start=2):
    cnes_library[excel_row] = {
        "State": str(row[0]).strip() if row[0] is not None else "",  # State in column A
        "City": str(row[1]).strip() if row[1] is not None else "",   # City in column B
        "CNES": str(row[2]).strip() if row[2] is not None else "",   # CNES in column C
        "Latitude": str(row[3]).strip() if row[3] is not None else "", # Latitude in column D
        "Longitude": str(row[4]).strip() if row[4] is not None else "", # Longitude in column E
        "REGIC Label": str(row[5]).strip() if row[5] is not None else "", # REGIC Label in column F
        "Address": str(row[6]).strip() if row[6] is not None else "", # Address in column G
        }
#print(list(cnes_library.items())[:5])

#Create Library for Relatório sheet
rel_library = {}
for excel_row, row in enumerate(rel_wb.iter_rows(min_row=2, values_only=True), start=2):
        rel_library[excel_row] = {
            "State": str(row[0]).strip() if row[0] is not None else "",  # State in column A
            "CNES": str(row[1]).strip() if row[1] is not None else "",   # CNES in column B
            "City": str(row[2]).strip() if row[2] is not None else "",   # City in column C
            "ZIP Code (sheet)": str(row[3]).strip() if row[3] is not None else "", # ZIP Code in column D
            "Address (sheet)": str(row[4]).strip() if row[4] is not None else "", # Address in column E
            "Latitude": str(row[6]).strip() if row[6] is not None else "", # Latitude in column G
            "Longitude": str(row[7]).strip() if row[7] is not None else "", # Longitude in column H
            "ZIP Code (site)": str(row[10]).strip() if row[10] is not None else "", # ZIP Code in column K
            "Address (site)": str(row[11]).strip() if row[11] is not None else "", # Address in column L
        }

#print(list(rel_library.items())[:5])

#Create CNES matches library
cnes_matches = {}
for cnes_row, cnes_data in cnes_library.items():
    for rel_row, rel_data in rel_library.items():
        if str(cnes_data["CNES"]) == str(rel_data["CNES"]):
            cnes_matches[cnes_row] = {
                "CNES": str(cnes_data["CNES"]) if cnes_data["CNES"] is not None else "",
                "State": str(cnes_data["State"]) if cnes_data["State"] is not None else "",
                "City": str(cnes_data["City"]) if cnes_data["City"] is not None else "",
                "ZIP Code (sheet)": str(rel_data["ZIP Code (sheet)"]) if rel_data["ZIP Code (sheet)"] is not None else "",
                "ZIP Code (site)": str(rel_data["ZIP Code (site)"]) if rel_data["ZIP Code (site)"] is not None else "",
                "Address (sheet)": str(rel_data["Address (sheet)"]) if rel_data["Address (sheet)"] is not None else "",
                "Address (site)": str(rel_data["Address (site)"]) if rel_data["Address (site)"] is not None else "",
                "Latitude (Relatorio)": str(rel_data["Latitude"]) if rel_data["Latitude"] is not None else "",
                "Longitude (Relatorio)": str(rel_data["Longitude"]) if rel_data["Longitude"] is not None else "",
                "Latitude (CNES)": str(cnes_data["Latitude"]) if cnes_data["Latitude"] is not None else "",
                "Longitude (CNES)": str(cnes_data["Longitude"]) if cnes_data["Longitude"] is not None else "",
                "REGIC Label": str(cnes_data["REGIC Label"]) if cnes_data["REGIC Label"] is not None else ""
            }
            break
#print(list(cnes_matches.items())[:5])
print("Number of CNES matches:", len(cnes_matches))

#Create library for CNES outliers
cnes_outliers_lib = {}
for cnes_row, cnes_data in cnes_library.items():
    if cnes_data["CNES"] not in [match["CNES"] for match in cnes_matches.values()]:
        cnes_outliers_lib[cnes_row] = cnes_data

#print(list(cnes_outliers_lib.items())[:5])
print("Number of CNES outliers:", len(cnes_outliers_lib))

#Create library for Relatório outliers
rel_outliers_lib = {}
for rel_row, rel_data in rel_library.items():
    if rel_data["CNES"] not in [match["CNES"] for match in cnes_matches.values()]:
        rel_outliers_lib[rel_row] = rel_data

#print(list(rel_outliers_lib.items())[:5])
print("Number of Relatório outliers:", len(rel_outliers_lib))

#Load libraries into the columns of the new excel file
for row, data in cnes_matches.items():
    new_wb.active.cell(row=row, column=1).value = data["CNES"]
    new_wb.active.cell(row=row, column=2).value = data["State"]
    new_wb.active.cell(row=row, column=3).value = data["City"]
    new_wb.active.cell(row=row, column=4).value = data["ZIP Code (sheet)"]
    new_wb.active.cell(row=row, column=5).value = data["ZIP Code (site)"]
    new_wb.active.cell(row=row, column=6).value = data["Address (sheet)"]
    new_wb.active.cell(row=row, column=7).value = data["Address (site)"]
    new_wb.active.cell(row=row, column=8).value = data["Latitude (Relatorio)"]
    new_wb.active.cell(row=row, column=9).value = data["Longitude (Relatorio)"]
    new_wb.active.cell(row=row, column=10).value = data["Latitude (CNES)"]
    new_wb.active.cell(row=row, column=11).value = data["Longitude (CNES)"]
    new_wb.active.cell(row=row, column=12).value = data["REGIC Label"]

for row, data in cnes_outliers_lib.items():
    new_wb.active.cell(row=row, column=13).value = data["CNES"]
    new_wb.active.cell(row=row, column=14).value = data["State"]
    new_wb.active.cell(row=row, column=15).value = data["City"]
    new_wb.active.cell(row=row, column=16).value = data["ZIP Code (sheet)"]
    new_wb.active.cell(row=row, column=17).value = data["ZIP Code (site)"]
    new_wb.active.cell(row=row, column=18).value = data["Address (sheet)"]
    new_wb.active.cell(row=row, column=19).value = data["Address (site)"]
    new_wb.active.cell(row=row, column=20).value = data["Latitude (Relatorio)"]
    new_wb.active.cell(row=row, column=21).value = data["Longitude (Relatorio)"]
    new_wb.active.cell(row=row, column=22).value = data["Latitude (CNES)"]
    new_wb.active.cell(row=row, column=23).value = data["Longitude (CNES)"]
    new_wb.active.cell(row=row, column=24).value = data["REGIC Label"]

for row, data in rel_outliers_lib.items():
    new_wb.active.cell(row=row, column=25).value = data["CNES"]
    new_wb.active.cell(row=row, column=26).value = data["State"]
    new_wb.active.cell(row=row, column=27).value = data["City"]
    new_wb.active.cell(row=row, column=28).value = data["ZIP Code (sheet)"]
    new_wb.active.cell(row=row, column=29).value = data["ZIP Code (site)"]
    new_wb.active.cell(row=row, column=30).value = data["Address (sheet)"]
    new_wb.active.cell(row=row, column=31).value = data["Address (site)"]
    new_wb.active.cell(row=row, column=32).value = data["Latitude (Relatorio)"]
    new_wb.active.cell(row=row, column=33).value = data["Longitude (Relatorio)"]
    new_wb.active.cell(row=row, column=34).value = data["Latitude (CNES)"]
    new_wb.active.cell(row=row, column=35).value = data["Longitude (CNES)"]
    new_wb.active.cell(row=row, column=36).value = data["REGIC Label"]