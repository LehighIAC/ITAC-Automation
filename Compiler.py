"""
Semi-automatically compile IAC report
Usage: Copy all ARs into the ARs folder, Update info in Info.json5 and Utility.json5,
then run this script.
"""


import json5, sys, os, locale
from docx import Document, shared
from python_docx_replace import docx_replace, docx_blocks
# Get the path of the current script, auxilliary functions are under Shared folder
script_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(script_path, 'Shared'))
from IAC import *
import pandas as pd
import datetime
from docx.enum.text import WD_ALIGN_PARAGRAPH

# If ARs/Sorted/ folder doesn't exist, create one
os.makedirs(os.path.join('ARs', 'Sorted'), exist_ok=True)

# If Energy Charts.fld doesn't exist, exit
if not os.path.exists(os.path.join(script_path, 'Energy Charts.fld')):
    print('Energy Charts.fld not found. Please save the energy chart as web page (.thm).')
    exit()

# Load config file and convert everything to local variables
iacDict = json5.load(open(os.path.join(script_path, 'Info.json5')))
iacDict.update(json5.load(open(os.path.join(script_path, 'Utility.json5'))))
locals().update(iacDict)

# Initialize dataframe
columns = ["isAAR", "File Name", "ARC No.", "Description", "Electricity (kWh)", "Electricity (MMBtu)", "Demand (kW)"
           , "Natural Gas (MMBtu)", "Other Energy Type", "Other Energy Amount", "Other Resource Type", "Other Resource Amount"
           , "Savings Type", "Savings Value", "Annual Cost Savings", "Implementation Cost", "Payback Period"]
df = pd.DataFrame(columns=columns)

# Set locale to en_US
locale.setlocale(locale.LC_ALL, 'en_US')

# Get all .docx files in the current directory and extract information
ARList = [f for f in os.listdir('ARs') if f.endswith('.docx')]
AR_id = 0
for ARdoc in ARList:
    doc = Document(os.path.join('ARs', ARdoc))
    ARinfo = {}
    # check if the document is an AAR
    ARinfo['isAAR'] = ("AAR" in doc.paragraphs[0].text.split(':')[0])
    # Record file name
    ARinfo['File Name'] = ARdoc
    # Parse the title of the .docx file
    ARinfo['Description'] = doc.paragraphs[0].text.split(':')[1].strip()
    # Read the first table in .docx files
    table = doc.tables[0]
    for row in table.rows:
        key = row.cells[0].text
        value = row.cells[1].text
        # Parse ARC Number
        if "arc" in key.lower() and "number" in key.lower():
            ARinfo['ARC No.'] = value
        # Parse Annual Cost Savings
        elif "annual" in key.lower() and "cost" in key.lower():
            # convert currency to interger
            ARinfo['Annual Cost Savings'] = locale.atof(value.strip("$"))
        # Parse Implementation Cost
        elif "implementation" in key.lower():
            # convert currency to interger
            ARinfo['Implementation Cost'] = locale.atof(value.strip("$"))
        # If Payback Period skip (Doesn't matter, will calculate later)
        elif "payback" in key.lower():
            continue
        # Parse Electricity
        elif "electricity" in key.lower():
            ARinfo['Electricity (kWh)'] = locale.atof(value.split(' ')[0])
        # Parse Demand
        elif "demand" in key.lower():
            ARinfo['Demand (kW)'] = locale.atof(value.split(' ')[0])
        # Parse Natural Gas
        elif "natural" in key.lower():
            ARinfo['Natural Gas (MMBtu)'] = locale.atof(value.split(' ')[0])
        # Parse undefined type
        else:
            # If the value contains mmbtu, parse it as other energy
            if "mmbtu" in value.lower():
                # Remove the last word (usually "savings")
                ARinfo['Other Energy Type'] = key.rsplit(' ', 1)[0]
                # Parse number
                ARinfo['Other Energy Amount'] = locale.atof(value.split(' ')[0])
            # If not, parse it as other resource
            else:
                # Remove the last word (usually "savings")
                ARinfo['Other Resource Type'] = key.rsplit(' ', 1)[0]
                # Keep the whole string
                ARinfo['Other Resource Amount'] = value   
    # Add dictionary to dataframe
    for key in ARinfo:
        df.loc[AR_id, key] = ARinfo[key]
    AR_id += 1

## Calculate on columns
# Calculate payback period
df['Payback Period'] = df['Implementation Cost'] / df['Annual Cost Savings']
# Convert Electricity to MMBtu
df['Electricity (MMBtu)'] = df['Electricity (kWh)']* 0.003413/0.33
# Sort df by payback period
df = df.sort_values(by=['Payback Period'])

## Format Savings strings
for index, row in df.iterrows():
    ST = ""
    SV = ""
    if pd.notna(row['Electricity (kWh)']):
        ST = ST + "Electricity" + '\n\n'
        SV = SV + locale.format_string('%d',row['Electricity (kWh)'], grouping=True) + ' kWh' + '\n'
        SV = SV + '(' + locale.format_string('%d',row['Electricity (MMBtu)'], grouping=True) + ' MMBtu)' + '\n'
    if pd.notna(row['Demand (kW)']):
        ST = ST + "Demand" + '\n'
        SV = SV + locale.format_string('%d',row['Demand (kW)'], grouping=True) + ' kW' + '\n'
    if pd.notna(row['Natural Gas (MMBtu)']):
        ST = ST + "Natural Gas" + '\n'
        SV = SV + locale.format_string('%d',row['Natural Gas (MMBtu)'], grouping=True)  + ' MMBtu' + '\n'
    if pd.notna(row['Other Energy Type']):
        ST = ST + row['Other Energy Type'] + '\n'
        SV = SV + locale.format_string('%d',row['Other Energy Amount'], grouping=True)  + ' MMBtu' '\n'
    if pd.notna(row['Other Resource Type']):
        ST = ST + row['Other Resource Type'] + '\n'
        SV = SV + row['Other Resource Amount'] + '\n'
    ST = ST.rstrip('\n')
    SV = SV.rstrip('\n')
    df.at[index, 'Savings Type'] = ST
    df.at[index, 'Savings Value'] = SV

## Summation statistics
# Split df to AR
AR_df = df[df['isAAR'] == False]
# reorder index
AR_df = AR_df.reset_index(drop=True)
# AR statistics
EkWh = AR_df['Electricity (kWh)'].sum(axis=0, skipna=True)
EMMBtu = AR_df['Electricity (MMBtu)'].sum(axis=0, skipna=True)
NMMBtu = AR_df['Natural Gas (MMBtu)'].sum(axis=0, skipna=True)
OMMBtu = AR_df['Other Energy Amount'].sum(axis=0, skipna=True)
# Add up all energy
ARMMBtu = round(EMMBtu + NMMBtu + OMMBtu)
# Calculate CO2
CO2 = round((53 * NMMBtu + 0.22 * EkWh)/1000)
# Add up all cost
ACS = AR_df['Annual Cost Savings'].sum(axis=0, skipna=True)
IC = AR_df['Implementation Cost'].sum(axis=0, skipna=True)
ARPB = round(IC / ACS, 1)
# Format as interger currency
locale._override_localeconv={'frac_digits':0}
iacDict['ARACS'] = locale.currency(ACS, grouping=True)
iacDict['ARIC'] = locale.currency(IC, grouping=True)


# Payback period string
iacDict['PB'] = payback(ACS, IC)

# Modify the title of the AR docx
for index, row in AR_df.iterrows():
    doc = Document(os.path.join('ARs', row['File Name']))
    # Change title and make it upper case
    doc.paragraphs[0].text = "AR "+ str(index+1) + ': ' + row['Description'].upper()
    # set font to bold
    doc.paragraphs[0].runs[0].bold = True
    # set font to upright
    doc.paragraphs[0].runs[0].italic = False
    # set font to no underline
    doc.paragraphs[0].runs[0].underline = False
    # set font size to 12
    doc.paragraphs[0].runs[0].font.size = shared.Pt(12)
    doc.save(os.path.join('ARs', 'Sorted', 'AR'+ str(index+1) + '.docx'))

# Check if there's at least 1 AAR
AAR = df['isAAR'].any()
if AAR:
    # Split df to AAR
    AAR_df = df[df['isAAR'] == True]
    # reorder index
    AAR_df = AAR_df.reset_index(drop=True)
    # AAR statistics
    EMMBtu = AAR_df['Electricity (MMBtu)'].sum(axis=0, skipna=True)
    NMMBtu = AAR_df['Natural Gas (MMBtu)'].sum(axis=0, skipna=True)
    OMMBtu = AAR_df['Other Energy Amount'].sum(axis=0, skipna=True)
    # Add up all energy
    AARMMBtu = round(EMMBtu + NMMBtu + OMMBtu)
    # Add up all cost
    ACS = AAR_df['Annual Cost Savings'].sum(axis=0, skipna=True)
    IC = AAR_df['Implementation Cost'].sum(axis=0, skipna=True)
    AARPB = round(IC / ACS, 1)
    # Format as interger currency
    iacDict['AARACS'] = locale.currency(ACS, grouping=True)
    iacDict['AARIC'] = locale.currency(IC, grouping=True)
    # Payback period number in the table

    # Modify the title of the AR docx
    for index, row in AAR_df.iterrows():
        doc = Document(os.path.join('ARs', row['File Name']))
        # Change title and make it upper case
        doc.paragraphs[0].text = "AAR "+ str(index+1) + ': ' + row['Description'].upper()
        # set font to bold
        doc.paragraphs[0].runs[0].bold = True
        # set font to upright
        doc.paragraphs[0].runs[0].italic = False
        # set font to no underline
        doc.paragraphs[0].runs[0].underline = False
        # set font size to 12
        doc.paragraphs[0].runs[0].font.size = shared.Pt(12)
        doc.save(os.path.join('ARs', 'Sorted', 'AAR'+ str(index+1) + '.docx'))

## Info.json5 Calculations
# Report date = today or 60 days after assessment, which ever is earlier
VD = datetime.datetime.strptime(VDATE, '%B %d, %Y')
RDATE = min(datetime.datetime.today(), VD + datetime.timedelta(days=60))
iacDict['RDATE'] = datetime.datetime.strftime(RDATE, '%B %d, %Y')

# Sort participant and contributor name list
PART=""
for name in PARTlist:
    # Sort by last name
    PARTlist.sort(key=lambda x: x.rsplit(' ', 1)[1])
    PART  = PART + name + '\n'
iacDict['PART'] = PART.rstrip('\n')
CONT=""
for name in CONTlist:
    # Sort by last name
    CONTlist.sort(key=lambda x: x.rsplit(' ', 1)[1])
    CONT  = CONT + name + '\n'
iacDict['CONT'] = CONT.rstrip('\n')

# Import docx template
doc = Document(os.path.join(script_path, 'IACtemplate.docx'))

# Add rows to AR table (Should be the 3rd table)
ARTable = doc.tables[2]
for index, row in AR_df.iterrows():
    ARrow = ARTable.rows[index+1].cells
    # Add ARC No.
    ARrow[0].text = 'AR ' + str(index+1) + '\n' + row['ARC No.']
    ARrow[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Add description
    ARrow[1].text = row['Description']
    ARrow[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    # Add savings type
    ARrow[2].text = row['Savings Type']
    ARrow[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Add savings value
    ARrow[3].text = row['Savings Value']
    ARrow[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Add annual cost savings
    ARrow[4].text = locale.currency(row['Annual Cost Savings'], grouping=True)
    ARrow[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # Add implementation cost
    ARrow[5].text = locale.currency(row['Implementation Cost'], grouping=True)
    ARrow[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # Add payback period
    ARrow[6].text = str(round(row['Payback Period'],1))
    ARrow[6].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # Set 3pt before and after paragraph
    for col in range(0,7):
        ARrow[col].paragraphs[0].paragraph_format.space_before = shared.Pt(3)
        ARrow[col].paragraphs[0].paragraph_format.space_after = shared.Pt(3)
# Delete unused rows (Currectly row 1-15 are empty)
for index in reversed(range(len(AR_df), 15)):
    ARTable._tbl.remove(ARTable.rows[index+1]._tr)

if AAR:
    # Add rows to AAR table (Should be the 4th table)
    AARTable = doc.tables[3]
    for index, row in AAR_df.iterrows():
        AARrow = AARTable.rows[index+1].cells
        # Add ARC No.
        AARrow[0].text = 'AAR ' + str(index+1) + '\n' + row['ARC No.']
        AARrow[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Add description
        AARrow[1].text = row['Description']
        AARrow[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        # Add savings type
        AARrow[2].text = row['Savings Type']
        AARrow[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Add savings value
        AARrow[3].text = row['Savings Value']
        AARrow[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Add annual cost savings
        AARrow[4].text = locale.currency(row['Annual Cost Savings'], grouping=True)
        AARrow[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # Add implementation cost
        AARrow[5].text = locale.currency(row['Implementation Cost'], grouping=True)
        AARrow[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # Add payback period
        AARrow[6].text = str(round(row['Payback Period'],1))
        AARrow[6].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # Set 3pt before and after paragraph
    for col in range(0,7):
        AARrow[col].paragraphs[0].paragraph_format.space_before = shared.Pt(3)
        AARrow[col].paragraphs[0].paragraph_format.space_after = shared.Pt(3)
    # Delete unused rows (Currectly row 1-5 are empty)
    for index in reversed(range(len(AAR_df), 5)):
        AARTable._tbl.remove(AARTable.rows[index+1]._tr)

# Add plant layout
# The file should be saved as "layout.png"
add_image(doc, '#LAYOUT', "layout.png", shared.Inches(6))

# Add energy chart images
add_image(doc, '#EUChart', os.path.join("Energy Charts.fld","image001.png"), shared.Inches(6))
add_image(doc, '#ECChart', os.path.join("Energy Charts.fld","image002.png"), shared.Inches(6))
add_image(doc, '#DUChart', os.path.join("Energy Charts.fld","image003.png"), shared.Inches(6))
add_image(doc, '#DCChart', os.path.join("Energy Charts.fld","image004.png"), shared.Inches(6))
add_image(doc, '#FUChart', os.path.join("Energy Charts.fld","image005.png"), shared.Inches(6))
add_image(doc, '#FCChart', os.path.join("Energy Charts.fld","image006.png"), shared.Inches(6))
add_image(doc, '#PieUChart', os.path.join("Energy Charts.fld","image007.png"), shared.Inches(6))
add_image(doc, '#PieCChart', os.path.join("Energy Charts.fld","image008.png"), shared.Inches(6))
add_image(doc, '#TotalChart', os.path.join("Energy Charts.fld","image009.png"), shared.Inches(9))

# Fill in energy chart tables from Energy Charts.xlsx
# Read electricity table from B6 to I19
edf = pd.read_excel("Energy Charts.xlsx", sheet_name="Raw Data", skiprows = 5, nrows=13, usecols = 'B:I')
# Read fuel table from K6 to N19
fdf = pd.read_excel("Energy Charts.xlsx", sheet_name="Raw Data", skiprows = 5, nrows=13, usecols = 'K:N')

# Add rows to electricity table (Should be the 10th table)
edfTable = doc.tables[9]
for index, row in edf.iterrows():
    edfrow = edfTable.rows[index+3].cells
    # Add Month
    edfrow[0].text = edf.iloc[(index, 0)]
    edfrow[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for col in range(1,8):
        # Add interger with thousand separator
        edfrow[col].text = locale.format_string('%d',edf.iloc[(index, col)], grouping=True)
        edfrow[col].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # Bold the last row
    if index == 12:
        for col in range(0,8):
            edfrow[col].paragraphs[0].runs[0].bold = True

# Add rows to fuel table (Should be the 11th table)
fdfTable = doc.tables[10]
for index, row in fdf.iterrows():
    fdfrow = fdfTable.rows[index+3].cells
    # Add Month
    fdfrow[0].text = fdf.iloc[(index, 0)]
    fdfrow[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for col in range(1,4):
        # Add interger with thousand separator
        fdfrow[col].text = locale.format_string('%d',fdf.iloc[(index, col)], grouping=True)
        fdfrow[col].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # Bold the last row
    if index == 12:
        for col in range(0,4):
            fdfrow[col].paragraphs[0].runs[0].bold = True

# Remove AAR blocks
docx_blocks(doc, AAR = AAR)

# Formatting
# Add all numbers in local variables to iacDict
iacDict.update({key: value for (key, value) in locals().items() if type(value) == int or type(value) == float})

# Format numbers to string with thousand separator
iacDict = grouping_num(iacDict)

# Format energy cost
locale._override_localeconv={'frac_digits':3}
iacDict['EC'] = locale.currency(EC, grouping=True)
locale._override_localeconv={'frac_digits':2}
iacDict['DC'] = locale.currency(DC, grouping=True)
iacDict['FC'] = locale.currency(FC, grouping=True)

# Replacing keys
docx_replace(doc, **iacDict)

filename = LE +'.docx'
doc.save(os.path.join(script_path, filename))

# Caveats
print("Please add Process Description, Major Equipment and Current Best Practices.")
print("Please copy and paste each sorted AR into the document.")
print("Please refresh ToC, tables and figures after running this script.")