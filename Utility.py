"""
Extract Data from Energy Charts.xlsx
Save statistics to Utility.json5
"""

import os, sys, locale, re, openpyxl
path = os.path.dirname(os.path.abspath(__file__))

# Read Energy Charts.xlsx
wb = openpyxl.load_workbook('Energy Charts.xlsx', data_only=True)
ws = wb['Raw Data']

## Save statistics to Utility.json5
# Get Electricity cost from cell D21
EC = round(ws.cell(row=21, column=4).value,3)
# Get Demand cost from cell D23
DC = round(ws.cell(row=23, column=4).value,2)
# Get Natural Gas cost from cell D24
NGC = round(ws.cell(row=24, column=4).value,2)
# Open Utility.json5 as text
with open('utility.json5', 'r') as f:
    utility = f.read()
# Replace values in Utility.json5
utility = re.sub(r'EC: \d+\.\d+', 'EC: ' + str(EC), utility)
utility = re.sub(r'DC: \d+\.\d+', 'DC: ' + str(DC), utility)
utility = re.sub(r'NGC: \d+\.\d+', 'NGC: ' + str(NGC), utility)
# Save utility.json5
with open('utility.json5', 'w') as f:
    f.write(utility)