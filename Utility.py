"""
Extract Data from Energy Charts.xlsx
Save statistics to Utility.json5
"""

import os, re, openpyxl
path = os.path.dirname(os.path.abspath(__file__))

# Read Energy Charts.xlsx
wb = openpyxl.load_workbook('Energy Charts.xlsx', data_only=True)
# Get Raw Data worksheet
ws = wb['Raw Data']

## Save statistics to Utility.json5
# Get Electricity cost from cell D21, 3 digits
EC = round(ws.cell(row=21, column=4).value,3)
# Get Demand cost from cell D23, 2 digits
DC = round(ws.cell(row=23, column=4).value,2)
# Get Fuel cost from cell D24, 2 digits
FC = round(ws.cell(row=24, column=4).value,2)
# Get Fuel type from cell Q2, string
Fuel = ws.cell(row=2, column=17).value
# Get Start Month from cell B7, string
StartMo = ws.cell(row=7, column=2).value.strftime("%b %y")
# Get End Month from cell B18, string
EndMo = ws.cell(row=18, column=2).value.strftime("%b %y")

# Get Total Electricity kWh from cell C19
TotalEkWh = round(ws.cell(row=19, column=3).value)
# Get Total Electricity MMBtu from cell I19
TotalEBtu = round(ws.cell(row=19, column=9).value)
# Get Total Demand kW from cell E19
TotalDkW = round(ws.cell(row=19, column=5).value)
# Get Total Fuel MMBtu from cell M19
TotalFBtu = round(ws.cell(row=19, column=13).value)

# Get Total Energy worksheet
ws = wb['Total Energy']
# Get Total Electricity cost from cell E5+E6
TotalECost = round(ws.cell(row=5, column=5).value+ws.cell(row=6, column=5).value)
# Get Total Fuel cost from cell E7
TotalFCost = round(ws.cell(row=7, column=5).value)
# Get Total Energy MMBtu from cell D8
TotalBtu = round(ws.cell(row=8, column=4).value)
# Get Total Energy Cost from cell E8
TotalCost = round(ws.cell(row=8, column=5).value)

# Write Natural Gas Cost for compatibility.
if Fuel == 'Natural Gas':
    NGC = FC
else:
    NGC = 0

# Open Utility.json5 as text
with open('utility.json5', 'r') as f:
    utility = f.read()

# Replace values in Utility.json5
utility = re.sub(r'EC: .*', 'EC: ' + str(EC) + ',', utility)
utility = re.sub(r'DC: .*', 'DC: ' + str(DC) + ',', utility)
utility = re.sub(r'FC: .*', 'FC: ' + str(FC) + ',', utility)
utility = re.sub(r'Fuel: .*', 'Fuel: "' + Fuel + '",', utility)
utility = re.sub(r'NGC: .*', 'NGC: ' + str(NGC) + ',', utility)
utility = re.sub(r'StartMo: .*', 'StartMo: "' + str(StartMo) + '",', utility)
utility = re.sub(r'EndMo: .*', 'EndMo: "' + str(EndMo) + '",', utility)
utility = re.sub(r'TotalEkWh: .*', 'TotalEkWh: ' + str(TotalEkWh) + ',', utility)
utility = re.sub(r'TotalEBtu: .*', 'TotalEBtu: ' + str(TotalEBtu) + ',', utility)
utility = re.sub(r'TotalDkW: .*', 'TotalDkW: ' + str(TotalDkW) + ',', utility)
utility = re.sub(r'TotalFBtu: .*', 'TotalFBtu: ' + str(TotalFBtu) + ',', utility)
utility = re.sub(r'TotalECost: .*', 'TotalECost: ' + str(TotalECost) + ',', utility)
utility = re.sub(r'TotalFCost: .*', 'TotalFCost: ' + str(TotalFCost) + ',', utility)
utility = re.sub(r'TotalBtu: .*', 'TotalBtu: ' + str(TotalBtu) + ',', utility)
utility = re.sub(r'TotalCost: .*', 'TotalCost: ' + str(TotalCost) + ',', utility)

# Save utility.json5
with open('utility.json5', 'w') as f:
    f.write(utility)