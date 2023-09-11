"""
This script is used to generate the IAC recommendation for Recover Exhaust Gas Heat.
"""

import json5, sys, os, locale
from docx import Document
from python_docx_replace import docx_replace, docx_blocks
# Get the path of the current script
script_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(script_path, '..', 'Shared'))
from IAC import *
import numpy as np

# Import docx template
doc = Document(os.path.join(script_path, 'Recover Exhaust Gas Heat.docx'))
# Load config file and convert everything to local variables
iacDict = json5.load(open(os.path.join(script_path, 'Exhaust Heat.json5')))
iacDict.update(json5.load(open(os.path.join(script_path, '..', 'plant.json5'))))
locals().update(iacDict)

# Interpolation
TrhoList = np.array([0, 10, 20, 30, 40, 50, 60, 70, 80, 100, 120, 140, 160, 180, 200, 250, 300, 350, 400, 450, 500, 600, 700, 800, 1000, 1200, 1400, 1600])
rhoList = np.array([0.0862, 0.0844, 0.0826, 0.081, 0.0793, 0.0778, 0.0763, 0.0749, 0.0735, 0.0709, 0.0685, 0.0662 , 0.0641, 0.0621\
                        , 0.0602, 0.0559, 0.0522, 0.0489, 0.0461, 0.0436, 0.041, 0.0371, 0.034, 0.0315, 0.0272, 0.0239, 0.0213, 0.0193])
RHO = round(np.interp(TI, TrhoList, rhoList).item(),3)

TCpList = np.array([-352, -318, -313, -280, -244, -208, -172, -136, -99.7, -63.7, -27.7, 8.3, 32, 44.3, 60, 80.3, 116, 152, 188, 224, 260, 440, 620, 800, 980, 1160, 1520, 2240, 2960])
CpList = np.array([0.2802, 0.251, 0.1791, 0.1739, 0.1726, 0.1716, 0.1713, 0.1712, 0.1711, 0.1711, 0.1711, 0.1712, 0.1713, 0.1713, 0.1714\
                   , 0.1715, 0.1718, 0.1721, 0.1725, 0.173, 0.1735, 0.1773, 0.1825, 0.1881, 0.1939, 0.1991, 0.2082, 0.2204, 0.2277])
CP = round(np.interp(TI, TCpList, CpList).item(),3)

# Calculations
NGS = round(CFM * RHO * 60 * CP * (TI - TO) * ETA * OH / 1e6)
CS = round(NGS * NGC)
ES = round(-HP * 0.746 * OH)
DS = round(-HP * 0.746 * 12)
ECS = round(ES * EC)
DCS = round(DS * DC)
PFC = ECS + DCS
ACS = CS + PFC

# Implementation
iacDict['PB'] = payback(ACS, IC)

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

# set the actual cost to integer
locale._override_localeconv={'frac_digits':0}
for cost in ['LR', 'CS', 'ECS', 'DCS', 'PFC', 'IC', 'ACS']:
    iacDict[cost] = locale.currency(eval(cost), grouping=True)

# Replacing keys
docx_replace(doc, **iacDict)

# Add equations
# Update numbers in local variables to formatted strings for easier access
locals().update(iacDict)
# Requires double backslash / curly bracket for LaTeX characters
NGSEqn = '\\frac{{ {0} \\times {1} \\times 60 \\times {2} \\times ({3} - {4}) \\times {5} \\times {6} }} {{ \\mathrm{{1,000,000}} }}' \
    .format(CFM, RHO, CP, TI, TO, ETA, OH)
add_eqn(doc, '#NGSEqn', NGSEqn)

# Save file as AR*.docx
filename = 'AR'+iacDict['AR']+'.docx'
doc.save(os.path.join(script_path, '..', 'ARs', filename))

# Caveats
print("Please manually change the font size of equations to 16.")
print("Please change implementation cost references if necessary.")