"""
This script is used to generate the IAC recommendation for Industrial fans to improve air circulation
"""

import json5, sys, os
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

## Calculations
# Operating hours
iac.OH = iac.HR * iac.DY * iac.WK

## Savings

## Rebate

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