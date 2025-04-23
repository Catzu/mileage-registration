# Installation
# pip install openpyxl
from openpyxl import Workbook, load_workbook
import os

# Creates excel file if it doesn't exist
if not os.path.exists("excel.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Data"
    wb.save("excel.xlsx")

# Load workbook and sheet
wk = load_workbook('excel.xlsx')
ws = wk.active

# Define header
header = ['Datum', 'postcode cafetaria', 'Postcode klant', 'adres klant( indien postcode vergeten te vragen)', 'aantal kilom heen en terug', 'vergoeding per km', 'totaalbedrag( niet invullen is formule)']

# Check if header exists in the first row
header_exists = False
for row in ws.iter_rows(values_only=True):
    if row == tuple(header):
        header_exists = True
        break

# Check first row and write header directly into row 1 if needed
first_row = [cell.value for cell in ws[1]]
if first_row != header:
    for col, val in enumerate(header, start=1):
        ws.cell(row=1, column=col, value=val)

# Get the next row number (where data will be added)
next_row = ws.max_row + 1 # +1 because we're going to append a new row

# Build the formula string with the correct row reference
formula = f"=E{next_row}*F{next_row}"

# Append the new row with the dynamic formula
ws.append(['18/04/2025', '5331 RD', '5324 JW', '', 1, 0.23, formula])

# Save workbook
wk.save('excel.xlsx')