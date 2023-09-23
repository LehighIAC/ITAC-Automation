"""
This script is used to generate the IAC recommendation for Switch to LED lighting.
"""

import json5, sys, os
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace, docx_blocks
sys.path.append('..') 
from Shared.IAC import *

# Load config file and convert everything to EasyDict
jsonDict = json5.load(open('Lighting.json5'))
jsonDict.update(json5.load(open(os.path.join('..', 'Utility.json5'))))
iac = EasyDict(jsonDict)

# Calculations
# Area 1
iac.ES1 = round((iac.CN1 * iac.CFW1 * iac.COH1 - iac.PN1 * iac.PFW1 * iac.POH1) / 1000.0)
iac.DS1 = round((iac.CN1 * iac.CFW1 - iac.PN1 * iac.PFW1) * iac.CF1 * 12.0 / 1000.0)
iac.BC1 = round(iac.PN1 * iac.BP1)
# Area 2
iac.ES2 = round((iac.CN2 * iac.CFW2 * iac.COH2 - iac.PN2 * iac.PFW2 * iac.POH2) / 1000.0)
iac.DS2 = round((iac.CN2 * iac.CFW2 - iac.PN2 * iac.PFW2) * iac.CF2 * 12.0 / 1000.0)
iac.BC2 = round(iac.PN2 * iac.BP2)
# Area 3
iac.ES3 = round((iac.CN3 * iac.CFW3 * iac.COH3 - iac.PN3 * iac.PFW3 * iac.POH3) / 1000.0)
iac.DS3 = round((iac.CN3 * iac.CFW3 - iac.PN3 * iac.PFW3) * iac.CF3 * 12.0 / 1000.0)
iac.BC3 = round(iac.PN3 * iac.BP3)
# Savings
iac.ES = iac.ES1 + iac.ES2 + iac.ES3
iac.DS = iac.DS1 + iac.DS2 + iac.DS3
iac.ECS = round(iac.ES * iac.EC)
iac.DCS = round(iac.DS * iac.DC)
iac.ACS = iac.ECS + iac.DCS
# Implementation
iac.MSC = iac.MSN * iac.MSPL
iac.BC = iac.BC1 + iac.BC2 + iac.BC3
iac.CN = iac.CN1 + iac.CN2 + iac.CN3
iac.LC = iac.BL1 * iac.CN1 + iac.BL2 * iac.CN2 + iac.BL3 * iac.CN3
iac.IC = iac.MSC + iac.BC + iac.LC
# Rebate
iac.RB = round(iac.ES * iac.RR)
iac.MRB = min(iac.RB, iac.IC/2)
iac.MIC = iac.IC - iac.MRB
iac.PB = payback(iac.ACS, iac.MIC)

# Combine words
AREAS = []
AREAS.append(iac.AREA1)
if iac.FLAG2:
    AREAS.append(iac.AREA2)
if iac.FLAG3:
    AREAS.append(iac.AREA3)
iac.AREAS = combine_words(AREAS)

# Motion sensors flag
if iac.MSN == 0:
    MSFLAG = False
else:
    MSFLAG = True

## Format strings
# set electricity cost / rebate to 3 digits accuracy
iac = dollar(['EC', 'RR'],iac,3)
# set the natural gas and demand to 2 digits accuracy
iac = dollar(['NGC', 'DC'],iac,2)
# set the rest to integer
varList = ['LR', 'MSPL', 'BL1', 'BL2', 'BL3', 'BP1', 'BP2', 'BP3', 'ECS', 'DCS', 'ACS', 'MSC', 'BC', 'LC', 'IC', 'RB', 'MRB', 'MIC']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('Switch to LED lighting.docx')

# Replacing keys
docx_replace(doc, **iac)

# Add equations
# Requires double backslash / curly bracket for LaTeX characters
ES1Eqn = '\\frac{{ {0} \\times {1} \\times {2} - {3} \\times {4} \\times {5} }} {{ \\mathrm{{1,000}} }}' \
    .format(iac.CN1, iac.CFW1, iac.COH1, iac.PN1, iac.PFW1, iac.POH1)
add_eqn(doc, '#ES1Eqn', ES1Eqn)

DS1Eqn = '\\frac{{ ({0} \\times {1} - {2} \\times {3}) \\times {4} \\times 12 }} {{ \\mathrm{{1,000}} }}' \
    .format(iac.CN1, iac.CFW1, iac.PN1, iac.PFW1, iac.CF1)
add_eqn(doc, '#DS1Eqn', DS1Eqn)

ES2Eqn = '\\frac{{ {0} \\times {1} \\times {2} - {3} \\times {4} \\times {5} }} {{ \\mathrm{{1,000}} }}' \
    .format(iac.CN2, iac.CFW2, iac.COH2, iac.PN2, iac.PFW2, iac.POH2)
add_eqn(doc, '#ES2Eqn', ES2Eqn)

DS2Eqn = '\\frac{{ ({0} \\times {1} - {2} \\times {3}) \\times {4} \\times 12 }} {{ \\mathrm{{1,000}} }}' \
    .format(iac.CN2, iac.CFW2, iac.PN2, iac.PFW2, iac.CF2)
add_eqn(doc, '#DS2Eqn', DS2Eqn)

ES3Eqn = '\\frac{{ {0} \\times {1} \\times {2} - {3} \\times {4} \\times {5} }} {{ \\mathrm{{1,000}} }}' \
    .format(iac.CN3, iac.CFW3, iac.COH3, iac.PN3, iac.PFW3, iac.POH3)
add_eqn(doc, '#ES3Eqn', ES3Eqn)

DS3Eqn = '\\frac{{ ({0} \\times {1} - {2} \\times {3}) \\times {4} \\times 12 }} {{ \\mathrm{{1,000}} }}' \
    .format(iac.CN3, iac.CFW3, iac.PN3, iac.PFW3, iac.CF3)
add_eqn(doc, '#DS3Eqn', DS3Eqn)

# Remove empty blocks
docx_blocks(doc, area1 = (iac.FLAG2 or iac.FLAG3))
docx_blocks(doc, area2 = iac.FLAG2)
docx_blocks(doc, area3 = iac.FLAG3)
docx_blocks(doc, ms = MSFLAG)

# Save file as AR*.docx
filename = 'AR'+iac.AR+'.docx'
doc.save(os.path.join('..', 'ARs', filename))

# Caveats
print("Please manually change the font size of equations to 16.")
print("Please change implementation cost references if necessary.")