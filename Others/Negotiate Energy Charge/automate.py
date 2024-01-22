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
# Determining Unit and Title
if iac.TYPE == "natural gas":
  iac.UNIT = "MMBtu"
  iac.TYPET = "Natural Gas"
elif iac.TYPE == "electricity":
  iac.UNIT = "kWh"
  iac.TYPET = "Electricity"
elif iac.TYPE == "propane":
  iac.UNIT = "MMBtu"
  iac.TYPET = "Propane"
elif iac.TYPE == "demand":
  iac.UNIT = "kW"
  iac.TYPET = "Demand"
# Determining site based on state
if iac.STATE == "PA":
  if iac.TYPE == "natural gas":
    iac.SITE = "https://www.pagasswitch.com"
  elif (iac.TYPE == "electricity" or iac.TYPE == "demand"):
    iac.SITE = "https://www.papowerswitch.com"
elif iac.STATE == "NJ":
  iac.SITE = "https://nj.gov/njpowerswitch/"
# Savings
iac.ACS = iac.EU * (iac.NGC - iac.PEC)

## Format strings
# set to 2 digits accuracy
iac = dollar(['NGC'],iac,2)
# set the rest to integer
varList = ['EU', 'PEC', 'ACS']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('template.docx')


# Replacing keys
docx_replace(doc, **iac)

filename = 'AR'+iac.AR+'.docx'
doc.save(os.path.join('..', '..', 'ARs', filename))