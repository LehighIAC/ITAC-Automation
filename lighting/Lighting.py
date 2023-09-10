"""
This script is used to generate the IAC recommendation for Switch to LED lighting.
"""

import json5, sys, os, locale
from docx import Document
from python_docx_replace import docx_replace, docx_blocks
# Get the path of the current script
script_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(script_path, '..', 'shared'))
from IAC import *

# Import docx template
doc = Document(os.path.join(script_path, 'Switch to LED lighting.docx'))
# Load config file and convert everything to local variables
iacDict = json5.load(open(os.path.join(script_path, 'Lighting.json5')))
iacDict.update(json5.load(open(os.path.join(script_path, '..', 'plant.json5'))))
locals().update(iacDict)

# Remove empty blocks
docx_blocks(doc, area2 = FLAG2)
docx_blocks(doc, area3 = FLAG3)

# Calculations
# Area 1
ES1 = round((CN1 * CFW1 * COH1 - PN1 * PFW1 * POH1) / 1000.0)
DS1 = round((CN1 * CFW1 - PN1 * PFW1) * CF1 * 12.0 / 1000.0)
BC1 = round(PN1 * BP1)
# Area 2
ES2 = round((CN2 * CFW2 * COH2 - PN2 * PFW2 * POH2) / 1000.0)
DS2 = round((CN2 * CFW2 - PN2 * PFW2) * CF2 * 12.0 / 1000.0)
BC2 = round(PN2 * BP2)
# Area 3
ES3 = round((CN3 * CFW3 * COH3 - PN3 * PFW3 * POH3) / 1000.0)
DS3 = round((CN3 * CFW3 - PN3 * PFW3) * CF3 * 12.0 / 1000.0)
BC3 = round(PN3 * BP3)
# Savings
ES = ES1 + ES2 + ES3
DS = DS1 + DS2 + DS3
ECS = round(ES * EC)
DCS = round(DS * DC)
ACS = ECS + DCS
# Implementation
MSC = MSN * MSPL
BC = BC1 + BC2 + BC3
CN = CN1 + CN2 + CN3
LC = BL * CN
IC = MSC + BC + LC
# Rebate
RB = round(ES * RR)
MRB = min(RB, IC/2)
MIC = IC - MRB
iacDict['PB'] = payback(ACS, MIC)

# Formatting
# Add all numbers in local variables to iacDict
iacDict.update({key: value for (key, value) in locals().items() if type(value) == int or type(value) == float})

# Format numbers to string with thousand separator
iacDict = grouping_num(iacDict)

# set locale to US
locale.setlocale(locale.LC_ALL, 'en_US')

# set 3 digits accuracy for electricity cost
locale._override_localeconv={'frac_digits':3}
iacDict['EC'] = locale.currency(EC, grouping=True)
iacDict['RR'] = locale.currency(RR, grouping=True)

# set the natural gas and demand to 2 digits accuracy
locale._override_localeconv={'frac_digits':2}
iacDict['NGC'] = locale.currency(NGC, grouping=True)
iacDict['DC'] = locale.currency(DC, grouping=True)

# set the rest to integer
locale._override_localeconv={'frac_digits':0}
for cost in ['LR', 'MSPL', 'BL', 'BP1', 'BP2', 'BP3', 'ECS', 'DCS', 'ACS', 'MSC', 'BC', 'LC', 'IC', 'RB', 'MRB', 'MIC']:
    iacDict[cost] = locale.currency(eval(cost), grouping=True)

# Replacing keys
docx_replace(doc, **iacDict)

# Add equations
# Update numbers in local variables to formatted strings for easier access
locals().update(iacDict)
# Requires double backslash / curly bracket for LaTeX characters
ES1Eqn = '\\frac{{ {0} \\times {1} \\times {2} - {3} \\times {4} \\times {5} }} {{ \\mathrm{{1,000}} }}' \
    .format(CN1, CFW1, COH1, PN1, PFW1, POH1)
add_eqn(doc, '#ES1Eqn', ES1Eqn)

DS1Eqn = '\\frac{{ ({0} \\times {1} - {2} \\times {3}) \\times {4} \\times 12 }} {{ \\mathrm{{1,000}} }}' \
    .format(CN1, CFW1, PN1, PFW1, CF1)
add_eqn(doc, '#DS1Eqn', DS1Eqn)

ES2Eqn = '\\frac{{ {0} \\times {1} \\times {2} - {3} \\times {4} \\times {5} }} {{ \\mathrm{{1,000}} }}' \
    .format(CN2, CFW2, COH2, PN2, PFW2, POH2)
add_eqn(doc, '#ES2Eqn', ES2Eqn)

DS2Eqn = '\\frac{{ ({0} \\times {1} - {2} \\times {3}) \\times {4} \\times 12 }} {{ \\mathrm{{1,000}} }}' \
    .format(CN2, CFW2, PN2, PFW2, CF2)
add_eqn(doc, '#DS2Eqn', DS2Eqn)

ES3Eqn = '\\frac{{ {0} \\times {1} \\times {2} - {3} \\times {4} \\times {5} }} {{ \\mathrm{{1,000}} }}' \
    .format(CN3, CFW3, COH3, PN3, PFW3, POH3)
add_eqn(doc, '#ES3Eqn', ES3Eqn)

DS3Eqn = '\\frac{{ ({0} \\times {1} - {2} \\times {3}) \\times {4} \\times 12 }} {{ \\mathrm{{1,000}} }}' \
    .format(CN3, CFW3, PN3, PFW3, CF3)
add_eqn(doc, '#DS3Eqn', DS3Eqn)

# Save file as AR*.docx
filename = 'AR'+iacDict['AR']+'.docx'
doc.save(os.path.join(script_path, '..', 'ARs', filename))

# Caveats
print("Please manually change the font size of equations to 16.")
if not (FLAG2 and FLAG3):
    print("Please manually remove zeroes in the document.")
print("Please change implementation cost references if necessary.")