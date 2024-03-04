"""
This script is used to generate the IAC recommendation for Install Air Curtain for Doorways
"""

import json5, sys, os, num2words
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace, docx_blocks
sys.path.append(os.path.join('..', '..'))
from Shared.IAC import *

# Load utility cost
jsonDict = json5.load(open(os.path.join('..', '..', 'Utility.json5')))
# Load database
jsonDict.update(json5.load(open('database.json5')))
# Convert to easydict
iac = EasyDict(jsonDict)

## Constants
# Correction Coefficient; kWh/MMBtu
C1 = 293
# Time correction; days per week
TC = 5
# Conversion constant; m/yr
C2 = 6
# Coincidence factor; %
CF = 100
# Conversion constant; KW/HP
C3 = 0.746

## Calculations
# Total Heat transfer
iac.HT = iac.HTF * iac.AMT
# Summer heat transfer
iac.SHT = round(iac.HT * (iac.DY / TC) * C1 * (iac.OT - iac.RT)) / iac.TDC
# Summer operating hours for HVAC
iac.OHS = round(iac.HRHV * iac.DY * iac.WK)
# Total horsepower
iac.HP = iac.HPF * iac.AMT
# Summer operating hours for air curtains
iac.OHAC = iac.HRAC * iac.DY * iac.WK
# Electricity usage of the air curtain system
iac.EU = round(iac.HP * C3 * iac.OHAC)
# Demand usage for the air curtain system
iac.DU = round(iac.HP * C3 * C2 * CF/100)

## Table
iac.AREA = iac.DW * iac.DH
iac.TOTALAREA = iac.AREA * iac.AMT
# Total # of doors
if type(iac.AMT) == list:
  iac.TOTALDOORS = sum(iac.AMT)
else:
  iac.TOTALDOORS = iac.AMT

## Savings
# Summer energy savings
iac.SES = round(iac.SHT * (iac.EF/100 - iac.EFES/100))
# Summer demand savings
iac.SDS = round((iac.SES/iac.OHS) * C2 * CF/100)
# Energy savings
iac.ES = iac.SES - iac.EU
# Demand savings
iac.DS = iac.SDS - iac.DU
# Energy cost savings
iac.ECS = iac.ES * iac.EC
# Demand cost savings
iac.DCS = iac.DS * iac.DC
# Annual cost savings
iac.ACS = iac.ECS + iac.DCS

## Implementation cost
iac.IC = (iac.COST * iac.AMT) + iac.LABOR
iac.PB = payback(iac.ACS, iac.IC)

## Number to words
iac.AMTSTR = num2words.num2words(iac.AMT)
iac.HRSTR = num2words.num2words(iac.HRAC)

## Format strings
# set electricity cost to 3 digits accuracy
iac = dollar(['EC'],iac,3)
# set the natural gas and demand to 2 digits accuracy
iac = dollar(['DC'],iac,2)
# set the rest to integer
varList = ['ACS', 'IC', 'COST', 'LABOR']
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