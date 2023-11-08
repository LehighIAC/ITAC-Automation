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

# Add equations
# Requires double backslash / curly bracket for LaTeX characters
CEUEqn = '\\frac{{ {0} \\times 0.746\\times {1}\\%\\times {2}}}{{ {3}\\% }}'.format(iac.HP, iac.LF, iac.OHE, iac.EX)
add_eqn(doc, iac, '${CEUEqn}', CEUEqn)

PEUEqn = '\\frac{{ {0} \\times 0.746\\times {1}\\%\\times {2}}}{{ {3}\\% }}'.format(iac.HP, iac.FRT, iac.OHP, iac.PR)
add_eqn(doc, iac, '${PEUEqn}', PEUEqn)

# Replacing keys
docx_replace(doc, **iac)

# Save file as AR*.docx
filename = 'AR'+str(iac.AR)+'.docx'
doc.save(os.path.join('..', 'ARs', filename))

# Caveats
caveat("Please manually change the font size of equations to 16.")
caveat("Please change implementation cost references if necessary.")