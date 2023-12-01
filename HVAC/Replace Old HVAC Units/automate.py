"""
This script is used to generate the IAC recommendation for Install programmable thermostats
"""

import json5, sys, os
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace, docx_blocks
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
# Maintendance Factor
if iac.FM == True:
  iac.M = 0.01
else:
  iac.M = 0.03
# Convert to numpy array
iac.TON = np.array(iac.TON)
iac.SIZE = iac.TON * 12000
iac.EERB = np.array(iac.EERB)
iac.AGE = np.array(iac.AGE)
iac.EERC = iac.EERB * (1 - iac.M) ** iac.AGE
iac.TTON = np
iac.TSIZE =

## Payback

iac.PB = payback(iac.ACS, iac.IC)
## Format strings
# set electricity cost to 3 digits accuracy
iac = dollar(['EC'],iac,3)
# set the natural gas and demand to 2 digits accuracy
iac = dollar(['NGC', 'DC'],iac,2)
# set the rest to integer
varList = ['LR', 'PT', 'LB', 'MC', 'IC', 'ECS', 'NGCS', 'ACS']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('template.docx')

# Replacing keys
docx_replace(doc, **iac)

# If both false
if iac.COOL == False and iac.HEAT == False:
    raise Exception("You need to have at least one section enabled")
# If both true
if iac.COOL and iac.HEAT:
    iac.DOUBLE = True
else:
    iac.DOUBLE = False
# Remove table row
summary = doc.tables[0]
if iac.COOL == False:
    summary._tbl.remove(summary.rows[3]._tr)
if iac.HEAT == False:
    summary._tbl.remove(summary.rows[4]._tr)

docx_blocks(doc, COOL = iac.COOL)
docx_blocks(doc, HEAT = iac.HEAT)
docx_blocks(doc, DOUBLE = iac.DOUBLE)

filename = 'AR'+iac.AR+'.docx'
doc.save(os.path.join('..', '..', 'ARs', filename))

# Caveats
caveat("Please change implementation cost references if necessary.")