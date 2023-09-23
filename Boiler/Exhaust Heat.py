"""
This script is used to generate the IAC recommendation for Recover Exhaust Gas Heat.
"""

import json5, sys, os
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace, docx_blocks
sys.path.append('..') 
from Shared.IAC import *
import numpy as np

# Load config file and convert everything to EasyDict
jsonDict = json5.load(open('Exhaust Heat.json5'))
jsonDict.update(json5.load(open(os.path.join('..', 'Utility.json5'))))
iac = EasyDict(jsonDict)

# Interpolation
TrhoList = np.array([0, 10, 20, 30, 40, 50, 60, 70, 80, 100, 120, 140, 160, 180, 200, 250, 300, 350, 400, 450, 500, 600, 700, 800, 1000, 1200, 1400, 1600])
rhoList = np.array([0.0862, 0.0844, 0.0826, 0.081, 0.0793, 0.0778, 0.0763, 0.0749, 0.0735, 0.0709, 0.0685, 0.0662 , 0.0641, 0.0621\
                        , 0.0602, 0.0559, 0.0522, 0.0489, 0.0461, 0.0436, 0.041, 0.0371, 0.034, 0.0315, 0.0272, 0.0239, 0.0213, 0.0193])
iac.RHO = round(np.interp(iac.TI, TrhoList, rhoList).item(),3)

TCpList = np.array([-352, -318, -313, -280, -244, -208, -172, -136, -99.7, -63.7, -27.7, 8.3, 32, 44.3, 60, 80.3, 116, 152, 188, 224, 260, 440, 620, 800, 980, 1160, 1520, 2240, 2960])
CpList = np.array([0.2802, 0.251, 0.1791, 0.1739, 0.1726, 0.1716, 0.1713, 0.1712, 0.1711, 0.1711, 0.1711, 0.1712, 0.1713, 0.1713, 0.1714\
                   , 0.1715, 0.1718, 0.1721, 0.1725, 0.173, 0.1735, 0.1773, 0.1825, 0.1881, 0.1939, 0.1991, 0.2082, 0.2204, 0.2277])
iac.CP = round(np.interp(iac.TI, TCpList, CpList).item(),3)

# Calculations
iac.NGS = round(iac.CFM * iac.RHO * 60 * iac.CP * (iac.TI - iac.TO) * iac.ETA * iac.OH / 1e6)
iac.CS = round(iac.NGS * iac.NGC)
iac.ES = round(-iac.HP * 0.746 * iac.OH)
iac.DS = round(-iac.HP * 0.746 * 12)
iac.ECS = round(iac.ES * iac.EC)
iac.DCS = round(iac.DS * iac.DC)
iac.PFC = iac.ECS + iac.DCS
iac.ACS = iac.CS + iac.PFC

# Implementation
iac.PB = payback(iac.ACS, iac.IC)

## Format strings
# set electricity cost to 3 digits accuracy
iac = dollar(['EC'],iac,3)
# set the natural gas and demand to 2 digits accuracy
iac = dollar(['NGC', 'DC'],iac,2)
# set the rest to integer
varList = ['LR', 'CS', 'ECS', 'DCS', 'PFC', 'IC', 'ACS']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('Recover Exhaust Gas Heat.docx')

# Replacing keys
docx_replace(doc, **iac)

# Add equations
# Requires double backslash / curly bracket for LaTeX characters
NGSEqn = '\\frac{{ {0} \\times {1} \\times 60 \\times {2} \\times ({3} - {4}) \\times {5} \\times {6} }} {{ \\mathrm{{1,000,000}} }}' \
    .format(iac.CFM, iac.RHO, iac.CP, iac.TI, iac.TO, iac.ETA, iac.OH)
add_eqn(doc, '#NGSEqn', NGSEqn)

# Save file as AR*.docx
filename = 'AR'+iac.AR+'.docx'
doc.save(os.path.join('..', 'ARs', filename))

# Caveats
print("Please manually change the font size of equations to 16.")
print("Please change implementation cost references if necessary.")