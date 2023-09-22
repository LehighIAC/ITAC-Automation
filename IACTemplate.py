"""
Semi-automated IAC template creator.
Usage: Copy all ARs into the ARs folder, Update info in Info.json5,
then run this script.
"""


import os, json5, re, math, easydict, openpyxl, locale
from docx import Document

# Get Info from Info.json5
script_path = os.path.dirname(os.path.abspath(__file__))
jsonDict = json5.load(open(os.path.join(script_path, 'Info.json5')))
jsonDict.update(json5.load(open(os.path.join(script_path, 'Utility.json5'))))
# Convert to EasyDict
iac = easydict.EasyDict(jsonDict)

# Read IAC template
wb = openpyxl.load_workbook('IACAssessmentTemplate.v2.1.xlsx')
# Open "General Info" sheet
print("Filling General Info:")
ws = wb['General Info']
ws['B2'] = int(iac.SIC)
ws['B3'] = int(iac.NAICS)
ws['B4'] = iac.SALE
ws['B5'] = iac.EMPL
ws['B6'] = iac.AREA
# Capitalize the first letter
ws['B7'] = iac.PROD[0].upper() + iac.PROD[1:]
ws['B8'] = iac.ANPR
# Parse Production Units
PU = iac.PRUN.lower()
if "unit" in PU:
    ws['B9'] = 1
elif "piece" in PU or "pcs" in PU:
    ws['B9'] = 2
elif "pound" in PU or "lbs" in PU:
    ws['B9'] = 3
elif "ton" in PU:
    ws['B9'] = 4
elif "barrel" in PU or "bbl" in PU:
    ws['B9'] = 5
elif "gallon" in PU or "gal" in PU:
    ws['B9'] = 6
elif "feet" in PU or "foot" in PU:
    ws['B9'] = 7
elif "bushel" in PU:
    ws['B9'] = 8
else:
    print("Error: Unknown Production Unit")
    exit()
# Parse Production Hours
OH = re.findall(r'\d+(?:\.\d+)?', iac.PROH)
OH = [float(i) for i in OH]
ws['B10'] = round(math.prod(OH))
print("Please manually fill B11:B16.")

# Open "Energy-Waste Info" sheet
print("Filling Energy-Waste Info:")
ws = wb['Energy-Waste Info']
# Read Energy Charts.xlsx
ecwb = openpyxl.load_workbook('Energy Charts.xlsx', data_only=True)
# Get Raw Data worksheet
ecws = ecwb['Raw Data']
# Make a list of corresponding cells to copy
wsList = ['B3', 'C3', 'B4', 'C4', 'C5']
ecList = ['C19', 'D19', 'E19', 'F19', 'G19']

FuelType = ecws['Q2'].value
if FuelType == 'Natural Gas':
    wsList.extend(['B6', 'C6'])
elif FuelType == 'Propane' or FuelType == 'Butane':
    wsList.extend(['B7', 'C7'])
elif FuelType == 'Fuel Oil #1':
    wsList.extend(['B8', 'C8'])
elif FuelType == 'Fuel Oil #2':
    wsList.extend(['B9', 'C9'])
elif FuelType == 'Fuel Oil #4':
    wsList.extend(['B10', 'C10'])
elif FuelType == 'Fuel Oil #6':
    wsList.extend(['B11', 'C11'])
elif FuelType == 'Coal':
    wsList.extend(['B12', 'C12'])
else:
    print("Unknown Fuel Type")
    exit()
ecList.extend(['M19', 'N19'])

# Replace values from the list
for i in range(len(wsList)):
    ws[wsList[i]].value = round(ecws[ecList[i]].value)
print("Please manually fill B13:B23.")

# Open "Energy-Waste Info" sheet
print("Filling Recommendation Info:")
ws = wb['Recommendation Info']

print("Please manually fill Recommendation Info for cross validation")

# Save as new file
wb.save('IACAssessmentTemplate.xlsx')