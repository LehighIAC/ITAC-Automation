"""
This script is used to generate the IAC recommendation for Install Solar Panels.
"""

import json5, sys, os, locale
from docx import Document
from python_docx_replace import docx_replace, docx_blocks
# Get the path of the current script
script_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(script_path, '..', 'shared'))
from IAC import *
import requests, datetime

# Import docx template
doc = Document(os.path.join(script_path, 'Install an Array of Solar Panels - NJ.docx'))
# Load config file and convert everything to local variables
iacDict = json5.load(open(os.path.join(script_path, 'Solar Panel NJ.json5')))
locals().update(iacDict)

# Calculations
# Avaialble space ft2
AS = round(RS * ASR / 100)
# Capacity kW
CAP = round(AS / 100)
# Approx. energy savings, kWh
AES = CAP * 1200

# PVWatts API
# The api key should be replaced after I gradudate
parameters = {
'format': 'json',
'api_key': 'bMgehoZeIcJNoYFh2KHbZFJw2X7ZYDn2z1SUdpNR',
'system_capacity': CAP,
'module_type': 0,
'losses': 14.08,
'array_type': 0,
'tilt': 20,
'azimuth': 180,
'address': ZIP,
}
response = requests.request('GET', 'https://developer.nrel.gov/api/pvwatts/v8.json', params=parameters)
PVresults = response.json()
ES = round(PVresults.get('outputs').get('ac_annual'))

ACSel = round(ES * EC)
credits = round(ES / 1000)
ACSsu = round(AMV * credits)
ACS = ACSel + ACSsu

# Implementation cost
IC = round(CAP * PPW * 1000)
ITC = round(IC * ITCR / 100)
MIC = IC - ITC
iacDict['PB'] = payback(ACS, MIC)
iacDict['CM'] = datetime.datetime.now().strftime('%B %Y')

# Formatting
# Add all numbers in local variables to iacDict
iacDict.update({key: value for (key, value) in locals().items() if type(value) == int or type(value) == float})

# Format numbers to string with thousand separator
iacDict = grouping_num(iacDict)

# set locale to US
locale.setlocale(locale.LC_ALL, 'en_US')

# set 3 digits accuracy for electricity cost
locale._override_localeconv={'frac_digits':3}
iacDict['EC'] = locale.currency(EC, grouping=True)

# set the natural gas and demand to 2 digits accuracy
locale._override_localeconv={'frac_digits':2}
iacDict['NGC'] = locale.currency(NGC, grouping=True)
iacDict['DC'] = locale.currency(DC, grouping=True)
iacDict['PPW'] = locale.currency(PPW, grouping=True)

# set the rest to integer
locale._override_localeconv={'frac_digits':0}
for cost in ['LR', 'MIC', 'IC', 'ITC', 'AMV', 'ACSel', 'ACSsu', 'ACS']:
    iacDict[cost] = locale.currency(eval(cost), grouping=True)

# Replacing keys
docx_replace(doc, **iacDict)

# This is an AAR by default
filename = 'AAR'+iacDict['AR']+'.docx'
doc.save(os.path.join(script_path, '..', 'ARs', filename))

# Caveats
print("Please check if the grabbed info is correct.")