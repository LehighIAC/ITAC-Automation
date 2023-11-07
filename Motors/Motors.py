"""
This script is used to generate the IAC recommendation for Installing VFD on Electric Motors
"""

import json5, sys, os
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace, docx_blocks
sys.path.append('..') 
from Shared.IAC import *

# Load config file and convert everything to EasyDict
jsonDict = json5.load(open('Motors.json5'))
jsonDict.update(json5.load(open(os.path.join('..', 'Utility.json5'))))
iac = EasyDict(jsonDict)

## Calculations
# Current Demand Usage
iac.CDU = round((iac.HP * 0.746 * 1.0) / 0.85)
# Current time weighted energy consumption for a given motor
iac.CEU = round(iac.CDU * iac.OHE)
# Proposed Demand Usage
iac.PDU = round((iac.HP * 0.746 * 0.73) / 0.85)
# Projected time weighted energy consumptions for a given motor
iac.PEU = round(iac.PDU * iac.OHP)

## Savings
# Annual Energy Savings
iac.ES = iac.CEU - iac.PEU
# Annual Demand Savings
iac.DS = (iac.CDU - iac.PDU) * 12 * 1.0
# Estimated Cost Savings
iac.ECS = round(iac.ES * iac.EC)
# Demand Cost Savings
iac.DCS = round(iac.DS * iac.DC)
# Total Cost Savings
iac.TCS = iac.ECS + iac.DCS
# Total Installation Cost
iac.IC = iac.NC + iac.EIC

## Rebate
iac.RB = round(iac.RR * iac.ES)
iac.MIC = iac.IC - iac.RB
iac.PB = payback(iac.TCS, iac.MIC)

## Format strings
# set electricity cost / rebate to 3 digits accuracy
iac = dollar(['EC', 'RR'],iac,3)
# set demand to 2 digits accuracy
iac = dollar(['DC'],iac,2)
# set the rest to integer
varList = ['TCS', 'MIC', 'ECS', 'DCS', 'NC', 'EIC', 'IC']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('Install VFD on Electric Motor Template.docx')

# # Add equations
# # Requires double backslash / curly bracket for LaTeX characters
# ES1Eqn = '\\frac{{ {0} \\times {1} \\times {2} - {3} \\times {4} \\times {5} }} {{ \\mathrm{{1,000}} }}' \
#     .format(iac.CN1, iac.CFW1, iac.COH1, iac.PN1, iac.PFW1, iac.POH1)
# add_eqn(doc, iac, '${ES1Eqn}', ES1Eqn)

# DS1Eqn = '\\frac{{ ({0} \\times {1} - {2} \\times {3}) \\times {4} \\times 12 }} {{ \\mathrm{{1,000}} }}' \
#     .format(iac.CN1, iac.CFW1, iac.PN1, iac.PFW1, iac.CF1)
# add_eqn(doc, iac, '${DS1Eqn}', DS1Eqn)

# ES2Eqn = '\\frac{{ {0} \\times {1} \\times {2} - {3} \\times {4} \\times {5} }} {{ \\mathrm{{1,000}} }}' \
#     .format(iac.CN2, iac.CFW2, iac.COH2, iac.PN2, iac.PFW2, iac.POH2)
# add_eqn(doc, iac, '${ES2Eqn}', ES2Eqn)

# DS2Eqn = '\\frac{{ ({0} \\times {1} - {2} \\times {3}) \\times {4} \\times 12 }} {{ \\mathrm{{1,000}} }}' \
#     .format(iac.CN2, iac.CFW2, iac.PN2, iac.PFW2, iac.CF2)
# add_eqn(doc, iac, '${DS2Eqn}', DS2Eqn)

# ES3Eqn = '\\frac{{ {0} \\times {1} \\times {2} - {3} \\times {4} \\times {5} }} {{ \\mathrm{{1,000}} }}' \
#     .format(iac.CN3, iac.CFW3, iac.COH3, iac.PN3, iac.PFW3, iac.POH3)
# add_eqn(doc, iac, '${ES3Eqn}', ES3Eqn)

# DS3Eqn = '\\frac{{ ({0} \\times {1} - {2} \\times {3}) \\times {4} \\times 12 }} {{ \\mathrm{{1,000}} }}' \
#     .format(iac.CN3, iac.CFW3, iac.PN3, iac.PFW3, iac.CF3)
# add_eqn(doc, iac, '${DS3Eqn}', DS3Eqn)

# Replacing keys
docx_replace(doc, **iac)

# Save file as AR*.docx
filename = 'AR'+str(iac.AR)+'.docx'
doc.save(os.path.join('..', 'ARs', filename))

# Caveats
caveat("Please manually change the font size of equations to 16.")
caveat("Please change implementation cost references if necessary.")