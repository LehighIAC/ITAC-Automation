"""
This script is used to generate the IAC recommendation for Install programmable thermostats
"""

import json5, sys, os
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace, docx_blocks
sys.path.append(os.path.join('..', '..'))
from Shared.IAC import *
from datetime import datetime

# Load utility cost
jsonDict = json5.load(open(os.path.join('..', '..', 'Utility.json5')))
# Load database
jsonDict.update(json5.load(open('database.json5')))
# Convert to easydict
iac = EasyDict(jsonDict)

## Calculations
# Current Month
iac.CM = datetime.now().strftime("%h") + " " +datetime.now().strftime("%Y")

iac.TYPET = title_case(iac.TYPE)
iac.TYPE = iac.TYPE.lower()

# Determining Unit, Supplier Type and Current Energy Cost
if iac.TYPE == "natural gas" or iac.TYPE == "propane":
  iac.UNIT = "MMBtu"
  iac.CEC = iac.FC
  iac.TYPES = iac.TYPE
elif iac.TYPE == "electricity":
  iac.UNIT = "kWh"
  iac.CEC = iac.EC
  iac.TYPES = iac.TYPE
elif iac.TYPE == "demand":
  iac.UNIT = "kW"
  iac.CEC = iac.DC
  iac.TYPES = "electricity"
else:
  raise Exception("Energy type is not supported.")

# Determining site based on state
if iac.STATE == "PA":
  if iac.TYPE == "natural gas":
    iac.SITE = "https://www.pagasswitch.com"
  elif (iac.TYPE == "electricity" or iac.TYPE == "demand"):
    iac.SITE = "https://www.papowerswitch.com"
elif iac.STATE == "NJ":
  iac.SITE = "https://nj.gov/njpowerswitch/"
else:
  raise Exception("State is not supported yet.")

if iac.CEC <= iac.PEC:
  raise Exception("Proposed energy cost is higher than current energy cost.")

# Savings
iac.ACS = iac.EU * (iac.CEC - iac.PEC)

## Format strings
if iac.TYPE == "electricity":
  iac = dollar(['CEC','PEC'],iac,3)
else:
  iac = dollar(['CEC','PEC'],iac,2)
# set the rest to integer
varList = ['ACS']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('template.docx')

# If uses propane, there's no website
if iac.TYPE == "propane":
  docx_blocks(doc, PROPANE = False)
else:
  docx_blocks(doc, PROPANE = True)

# Replacing keys
docx_replace(doc, **iac)

savefile(doc, iac.AR)