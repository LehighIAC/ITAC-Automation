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
from meteostat import Point, Monthly, units

# Load utility cost
jsonDict = json5.load(open(os.path.join('..', '..', 'Utility.json5')))
# Load database
jsonDict.update(json5.load(open('database.json5')))
# Convert to easydict
iac = EasyDict(jsonDict)

## Retreiving AVerage Outside Temperature
# Use US zipcodes
nomi = pgeocode.Nominatim('us')
query = nomi.query_postal_code(iac.ZIP)
# Return lat and lon information
lat = query['latitude']
lon = query['longitude']

start = datetime(2018, 10, 1)
end = datetime(2022, 5, 31)
# Use coordinates for weather info
point = Point(lat, lon)
# Get monthly data
data = Monthly(point, start, end)
df = data.fetch()

# Extract average temperature
df = df['tavg'].dropna()
# Convert to degF
df = df * 9/5 + 32
iac.TO = round(df.mean())

## VFD table
Load = np.linspace(20, 100, num=17)
VFD = np.array([25, 28, 33, 38, 42, 47, 52, 57, 61, 65, 70, 75, 80, 85, 90, 95, 105])

# Power Fraction Caclulation
# Blow Off
if iac.CT == 1:
    iac.FPC = 100
    iac.CT = "blow off"
# Modulation
elif iac.CT == 2:
    iac.FPC = round(0.3 * iac.LF + 70)
    iac.CT = "modulation"
# Load/Unload
elif iac.CT == 3:
    iac.FPC = round(0.5 * iac.LF + 50)
    iac.CT = "load/unload"
# VFD
elif iac.CT == 4:
    iac.FPC = round(np.interp(iac.LF, Load, VFD).item())
else:
    raise Exception("Wrong control type!")

## Calculations
# Compressor Work Reduction
iac.CWR = round((iac.DT)/(iac.TI + 460) * 100, 2)
# Operating hours
iac.OH = iac.HR * iac.DY * iac.WK
# Power Reduction
iac.PR = round((iac.HP * 0.746 * (iac.FPC/100) * (iac.CWR/100)) / (iac.ETA/100), 1)

## Savings
# Electrcity Savings
iac.ES = round(iac.PR * iac.OH)
# Demand Savings
iac.DS = round(iac.PR * (iac.CF/100) * 12)
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
iac = dollar(['EC','DC'],iac,2)
# set the rest to integer
varList = ['ACS', 'ECS', 'DCS', 'IC']
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