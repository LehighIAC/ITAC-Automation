"""
Extract Data from Energy Charts.xlsx
Save statistics to Utility.json5
"""

import os, re, openpyxl

# Read Energy Charts.xlsx
wb = openpyxl.load_workbook(os.path.join('Energy Charts' ,'Energy Charts.xlsx'), data_only=True)
# Get Raw Data worksheet
ws = wb['Raw Data']

## Save statistics to Utility.json5
# Get Electricity cost from cell D21, 3 digits
EC = round(ws['D21'].value,3)
# Get Demand cost from cell D23, 2 digits
DC = round(ws['D23'].value,2)
# Get Fuel cost from cell D24, 2 digits
FC = round(ws['D24'].value,2)
# Get Fuel type from cell Q2, string
FuelType = ws['Q2'].value
# Get Fuel unit from cell Q3, string
FuelUnit = ws['Q3'].value
# Get Start Month from cell B7, string
StartMo = ws['B7'].value
# Get End Month from cell B18, string
EndMo = ws['B18'].value
# Get Total Electricity kWh from cell C19
TotalEkWh = round(ws['C19'].value)
# Get Total Electricity MMBtu from cell I19
TotalEBtu = round(ws['I19'].value)
# Get Total Demand kW from cell E19
TotalDkW = round(ws['E19'].value)
# Get Total Fuel MMBtu from cell M19
TotalFBtu = round(ws['M19'].value)

# Get Total Energy worksheet
ws = wb['Total Energy']
# Get Total Electricity cost from cell E5+E6
TotalECost = round(ws['E5'].value+ws['E6'].value)
# Get Total Fuel cost from cell E7
TotalFCost = round(ws['E7'].value)
# Get Total Energy MMBtu from cell D8
TotalBtu = round(ws['D8'].value)
# Get Total Energy Cost from cell E8
TotalCost = round(ws['E8'].value)

# Write Natural Gas Cost for compatibility.
if FuelType == 'Natural Gas':
    NGC = FC
else:
    NGC = 0

# Open Utility.json5 as text
try:
    with open('Utility.json5', 'r') as f:
        utility = f.read()
        f.close()
except FileNotFoundError:
    raise Exception('Utility.json5 not found.')

# Replace values in Utility.json5
utility = re.sub(r'EC: .*', 'EC: ' + str(EC) + ',', utility)
utility = re.sub(r'DC: .*', 'DC: ' + str(DC) + ',', utility)
utility = re.sub(r'FC: .*', 'FC: ' + str(FC) + ',', utility)
utility = re.sub(r'FuelType: .*', 'FuelType: "' + FuelType + '",', utility)
utility = re.sub(r'FuelUnit: .*', 'FuelUnit: "' + FuelUnit + '",', utility)
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

# Save Utility.json5
try:
    with open('Utility.json5', 'w') as f:
        f.write(utility)
        f.close()
except FileNotFoundError:
    raise Exception('Utility.json5 not found.')