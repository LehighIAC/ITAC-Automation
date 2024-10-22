"""
This script is used to generate the IAC recommendation for Installing VFD on Electric Motors
"""

import json5, sys, os
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace
sys.path.append(os.path.join('..', '..')) 
from Shared.IAC import *
import numpy as np

# Load utility cost
jsonDict = json5.load(open(os.path.join('..', '..', 'Utility.json5')))
# Load database
jsonDict.update(json5.load(open('database.json5')))
# Convert to easydict
iac = EasyDict(jsonDict)

## Calculations
# Operating hours
iac.OH = iac.HR * iac.DY * iac.WK

## Savings
# Annual Energy Savings
iac.ES = round(iac.HP * 0.746 * iac.LF/100 * iac.OH * ((1.5/100)/(iac.ETA/100)))
# Annual Demand Savings
iac.DS = round(iac.HP * 0.746 * iac.LF/100 * iac.CF/100 * 12 * ((1.5/100)/(iac.ETA/100)))
# Estimated Cost Savings
iac.ECS = round(iac.ES * iac.EC)
# Demand Cost Savings
iac.DCS = round(iac.DS * iac.DS)
# Total Cost Savings
iac.ACS = iac.ECS + iac.DCS
# Total Installation Cost
iac.IC = iac.CBELT * iac.AMT

## Rebate
iac.PB = payback(iac.ACS, iac.IC)

## Format strings
# set electricity cost / rebate to 3 digits accuracy
iac = dollar(['EC'],iac,3)
# set demand to 2 digits accuracy
iac = dollar(['DC'],iac,2)
# set the rest to integer
varList = ['ACS', 'ECS', 'DCS', 'CBELT', 'IC']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('template.docx')

# Replacing keys
docx_replace(doc, **iac)

savefile(doc, iac.REC)

# Caveats
caveat("Please change implementation cost references if necessary.")