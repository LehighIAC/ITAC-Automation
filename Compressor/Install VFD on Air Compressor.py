"""
This script is used to generate the IAC recommendation for Installing VFD on Electric Motors
"""

import json5, sys, os
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace, docx_blocks
sys.path.append('..') 
from Shared.IAC import *
import numpy as np

# Load config file and convert everything to EasyDict
jsonDict = json5.load(open('Install VFD on Air Compressor.json5'))
jsonDict.update(json5.load(open(os.path.join('..', 'Utility.json5'))))
iac = EasyDict(jsonDict)

## VFD table
Load = np.linspace(20, 100, num=17)
VFD = np.array([5, 6, 8, 11, 14, 17, 21, 26, 32, 38, 44, 50, 57, 64, 73, 86, 105])

## Calculations
# Operating hours
iac.OH = iac.HR * iac.DY * iac.WK
# Power FRaction with VFD
iac.FR = round(np.interp(iac.LF, Load, VFD).item())
# Current Power Draw
iac.CPD = round((iac.HP * 0.746) / (iac.ETAE/100))
# Proposed Power Draw
iac.PPD = round((iac.HP * 0.746 * (iac.FR/100)) / (iac.ETAP/100))

## Savings
# Annual Energy Savings
iac.ES = (iac.CPD - iac.PPD) * iac.OH
# Annual Demand Savings
iac.DS = (iac.CPD - iac.PPD) * (iac.CF/100) * 12
# Estimated Cost Savings
iac.ECS = round(iac.ES * iac.EC)
# Demand Cost Savings
iac.DCS = round(iac.DS * iac.DC)
# Total Cost Savings
iac.ACS = iac.ECS + iac.DCS
# Total Installation Cost
iac.IC = iac.VFD + iac.AIC

## Rebate
iac.RB = round(iac.RR * iac.ES)
iac.MRB = max(iac.RB, iac.IC/2)
iac.MIC = iac.IC - iac.RB
iac.PB = payback(iac.ACS, iac.MIC)

## Format strings
# set electricity cost / rebate to 3 digits accuracy
iac = dollar(['EC', 'RR'],iac,3)
# set demand to 2 digits accuracy
iac = dollar(['DC'],iac,2)
# set the rest to integer
varList = ['ACS', 'ECS', 'DCS', 'VFD', 'AIC', 'IC', 'RB', 'MRB', 'MIC']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('Install VFD on Electric Motor.docx')

# Replacing keys
docx_replace(doc, **iac)

# Save file as AR*.docx
filename = 'AR'+str(iac.AR)+'.docx'
doc.save(os.path.join('..', 'ARs', filename))

# Caveats
caveat("Please change implementation cost references if necessary.")