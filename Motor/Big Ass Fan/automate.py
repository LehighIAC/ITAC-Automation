"""
This script is used to generate the IAC recommendation for Industrial fans to improve air circulation
"""

import json5, sys, os, num2words
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace
sys.path.append(os.path.join('..', '..')) 
from Shared.IAC import *

# Load utility cost
jsonDict = json5.load(open(os.path.join('..', '..', 'Utility.json5')))
# Load database
jsonDict.update(json5.load(open('database.json5')))
# Convert to easydict
iac = EasyDict(jsonDict)

## Constants
# Conversion constant
C1 = 0.7457
# Coincidence factor, %
CF = 100

## Calculations
# Operating hours
iac.OH = iac.HR * iac.DY * iac.WK

## Savings
# Natural Gas Savings
iac.NGS = iac.PR/100 * iac.NGU
# Extra electricity consumption
iac.ES = round(-(iac.FAN) * iac.HP * C1 * iac.OH)
# Extra demand consumption
iac.DS = round(-(iac.FAN) * iac.HP * C1 * 6 * CF/100, 1)
# Natual Gas cost savings
iac.NGCS = round(iac.NGS * iac.NGC)
# Electricity cost savings
iac.ECS = round(iac.ES * iac.EC)
# Demand cost savings
iac.DCS = round(iac.DS * iac.DC)
# Annual cost savings
iac.ACS = iac.NGCS + iac.ECS + iac.DCS

## Rebate
# Total fan cost
iac.IC = iac.FAN * iac.COST
iac.PB = payback(iac.ACS, iac.IC)

## Format strings
# Convert to word
iac.FANStr = num2words.num2words(iac.FAN)
# set electricity cost / rebate to 3 digits accuracy
iac = dollar(['EC'],iac,3)
# set demand to 2 digits accuracy
iac = dollar(['DC', 'NGC'],iac,2)
# set the rest to integer
varList = ['NGCS', 'ECS', 'DCS', 'ACS', 'COST', 'IC']
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