"""
This script is used to generate the IAC recommendation for Reduce Compressor Set Pressure.
"""

import json5, sys, os
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace, docx_blocks
sys.path.append('..')
from Shared.IAC import *

# Load config file and convert everything to EasyDict
jsonDict = json5.load(open('Reduce Compressor Set Pressure.json5'))
jsonDict.update(json5.load(open(os.path.join('..', 'Utility.json5'))))
iac = EasyDict(jsonDict)

# Constants
AP = 14.7
k = 1.4

# Calculating Power reduction
iac.POW = 1-(((iac.RCP+AP)/AP)**((k-1)/(k*iac.N))-1)/(((iac.CCP+AP)/AP)**((k-1)/(k*iac.N))-1)
# 1 decimal point percent
iac.POW = round(iac.POW*100,1)

# Opearting Hours
iac.OH = iac.HR * iac.DY * iac.WK
# energy savings
iac.ES = int(iac.HP * iac.OH * iac.LF * iac.RF * 0.746 * (iac.POW / 100) / iac.ETA)

# demand saving
iac.DS = int(iac.ES * iac.CF * 12 / iac.OH)

# electricity cost savings
iac.ECS = int(iac.ES * iac.EC)
# demand cost savings
iac.DCS = int(iac.DS * iac.DC)

iac.ACS = iac.ECS + iac.DCS

iac.PB  = payback(iac.ACS, iac.IC)

## Format strings
# set electricity cost to 3 digits accuracy
iac = dollar(['EC'],iac,3)
# set demand to 2 digits accuracy
iac = dollar(['DC'],iac,2)
# set the rest to integer
varList = ['IC', 'ACS', 'ECS', 'DCS']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('Reduce Compressor Set Pressure.docx')

# Add equations
# Requires double backslash / curly bracket for LaTeX characters
POWEqn = '\\frac{{ \\left( \\frac{{ {0} + 14.7}}{{14.7}} \\right)^{{\\left( \\frac{{1.4-1}}{{1.4\\times {1} }} \\right)}}-1}}{{ \\left( \\frac{{ {2} + 14.7}}{{14.7}} \\right)^{{\\left( \\frac{{1.4-1}}{{1.4\\times {3} }} \\right)}}-1}}' \
    .format(iac.RCP, iac.N, iac.CCP, iac.N)
add_eqn(doc, iac, '${POWEqn}', POWEqn)

# Replacing keys
docx_replace(doc, **iac)

filename = 'AR'+iac.AR+'.docx'
doc.save(os.path.join('..', 'ARs', filename))

# Caveats
caveat("Please manually change the font size of equations to 16.")
caveat("Please change implementation cost references if necessary.")