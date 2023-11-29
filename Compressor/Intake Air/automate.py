"""
This script is used to generate the IAC recommendation for Air Intake Compressors
"""

import json5, sys, os
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace, docx_blocks
sys.path.append(os.path.join('..', '..')) 
from Shared.IAC import *
import numpy as np
import pgeocode
from datetime import datetime
from meteostat import Point, Monthly

# Load utility cost
jsonDict = json5.load(open(os.path.join('..', '..', 'Utility.json5')))
# Load database
jsonDict.update(json5.load(open('database.json5')))
# Convert to easydict
iac = EasyDict(jsonDict)

## Calculations
# Compressor Work Reduction
iac.CWR = int((iac.IT - iac.OT)/(iac.IT + 460)* 10000)
iac.CWR = float(iac.CWR/100)
# Operating hours
iac.OH = iac.HR * iac.DY * iac.WK
# Power Reduction
iac.PR = int((iac.HP * 0.746 * iac.LF * iac.CWR)/iac.EHR)
iac.PR = float(iac.PR/100)

## Retreiving AVerage Outside Temperature
# Turn zipcode into a string
zipcode_num = str(iac.OT)
# Use US zipcodes
nomi = pgeocode.Nominatim('us')
query = nomi.query_postal_code(zipcode_num)
# Return lat and lon information
lat = query['latitude']
lon = query['longitude']
# Start searching from Oct 2018 to May 2022
start = datetime(2018, 10, 1)
end = datetime(2022, 5, 1)
# Use coordinates for weather info
place = (lon, lat)

from urllib.parse import quote

# Assuming place, start, and end are defined elsewhere in your code
coordinates = f'({lon} {lat})'
url = f'/v2/monthly/{quote(coordinates)}.csv.gz'
data = Monthly(place, start, end)
data = data.fetch(url)

## Savings
# Electrcity Savings
iac.ES = round(iac.PR * iac.OH)
# Demand Savings
iac.DS = round(iac.PR * iac.CF/100 * iac.CC)
# Electrcicity Cost Savings
iac.ECS = round(iac.ES * iac.EC) 
# Demand Cost Savings
iac.DCS = round(iac.DS * iac.DC)
# Annual Cost Savigns
iac.ACS = round(iac.ECS + iac.DCS)

## Rebate
iac.PB = payback(iac.ACS, iac.IC)

## Format strings
# set to 2 digits accuracy
iac = dollar(['NGC'],iac,2)
# set the rest to integer
varList = ['ACS', 'ECS', 'DCS', 'IC']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('template.docx')

# Replacing keys
docx_replace(doc, **iac)

# Save file as AR*.docx
filename = 'AR'+str(iac.AR)+'.docx'
doc.save(os.path.join('..', '..', 'ARs', filename))

# Caveats
caveat("Please change implementation cost references if necessary.")