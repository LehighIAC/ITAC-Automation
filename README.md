# IAC-Automation
Automated Python script for Lehigh University Industrial Assessment Center
## Guide for New IAC Members
Of course, you need to have Microsoft Office installed.
### Windows
0. Install [VS code](https://code.visualstudio.com/download), [Anaconda](https://www.anaconda.com/download), [Github Desktop](https://desktop.github.com), and [git](https://gitforwindows.org/) (all default).
### macOS
0. Install [homebrew](https://brew.sh) then```brew install --cask visual-studio-code anaconda github```
### Linux
LibreOffice compatibility is not guaranteed.
### For all operating systems
1. In VS Code, go to `Extensions`(Ctrl+Shift+X), install `Python` (from Microsoft) and `JSON5 syntax` from mrmlnc.

2. Register a GitHub account with your Lehigh email, then **fork** the [main repository](https://github.com/LehighIAC/IAC-Automation/tree/main). It will make a copy under your account.

3. Sign in GitHub Desktop, **clone your fork** (not the main repository) to the local computer. The fork should be under your username.

4. **IMPORTANT**  Switch to **develop** branch.

4. After validating your proposed changes, go to Github Desktop to `commit` and `push` new code to your fork. *Remember to write detailed comments so other people can understand your proposed changes.*
   
5. Go to GitHub website (there should be a shortcut in Github Desktop) and send a pull request. *Remember to write detailed comments so other people can understand your proposed changes.*

6. DO NOT include any sensitive information such as plant name and address when contributing code.

7. After reviewing the code, The IAC Admin can approve and merge your proposed changes.

## Setting up Python environment
### Open Anaconda Terminal, or VS Code Terminal.
```
conda create -n iac python=3.8 
conda activate iac 
```
### Install the Following Packages
```
conda install json5 numpy pandas"<2" openpyxl requests
conda install -c conda-forge python-docx docxcompose easydict latex2mathml num2words pgeocode
pip install python-docx-replace meteostat
```
`conda` always has the highest priority. If not available, install packages from `conda-forge`. Don't install from `pip` unless you have to, otherwise there might be dependency issue.
### Configure VS Code Environment
In VS Code, press Ctrl+Shift+P, search `Python: Select Interpreter` and select the `iac` environment you just created.
### NOTE: IF YOU WISH TO REMOVE THIS ENVIRONMENT
```
conda remove --name iac --all
```
## Using templates
Is it suggested to work on a copy of this reposiotry when generating an IAC report
For your convenience, download the main branch from this link: https://codeload.github.com/LehighIAC/IAC-Automation/zip/refs/heads/main
### Energy Charts
1. Edit `Energy Charts.xlsx`. Select `fuel type` ,`fuel unit` and `start month`, then edit raw data (if copying from other spreadsheet, copy values only). The formatting is fully automatic and shouldn't be touched.
2. Save the workbook as `Web Page (.htm)` format in the same directory. DO NOT change the filename, all images will be  kept in `Energy Charts.fld` folder. Currently  this is the only stable way to save all charts as images.
3. Run `Utility.py` to extract energy usage data from the spreadsheet.
### Assessment Recommendations
1. Edit `.json5` database of any specific AR. Make sure the data type is matching the description.
2. Run the corresponding `.py` file. The output will be saved in `ARs` directory. Follow the instructions of the script if there's anything you need to adjust manually.

### Requirements of AR Files:
1. No requirement for filename, as long as it's `.docx`
2. Doesn't matter if the file is made from Python template, Excel template, or by hand. Please **break links** if you used Excel templates.
3. The title text should always be "AR x: Title" or "AAR x: Title". Case insensitive. In outline view the title should always be **Level 1**.
4. If there's any other type of energy savings, the unit should be `MMBtu`.
5. All subtitles, such as "Recommend Actions", "Anticipated Savings" should always be **body text** in outline view. Otherwise the automatic table of contents might be broken.
6. In rare cases, some .docx files are actually wrapped legacy .doc file (pre Word 2000) and is not supported by this tool. Please copy and paste everything to a new blank .docx file. 

### Compiling Report
1. Fill required plant information in `Compiler.json5`.
2. Copy all AR files(if you have any from other sources) into `ARs` directory.
3. Run `Compiler.py` to compile the final report.
4. Ctrl+A then F9 to refresh ToC, tables and figures, you may need to do it twice.
5. Fill the rest of the information manually.

## Supported AR Templates

### Boiler
* Recover Exhaust Gas Heat

### Compressor
* Draw Compressor Intake Air from Outside
* Use Compressor Exhaust to Heat during Winter Months
* Reduce Compressor Set Pressure
* Repair Leaks in Compressed Air Lines
* Install VFD on Air Compressor (Single Motor)

### HVAC
* Programmable Thermostat (based on degree hours)
* Replace Old HVAC Units

### Lighting
* Switch to LED lighting (supports any number of areas)

### Motors
* Install VFD on Electric Motor (Single Motor)

### Others
* Install Solar Panel (fully automated using PVWatts API)


## Developing New Templates
**Always make changes in the `develop` branch**!

An automated template is usually made of 3 parts:
1. `template.docx` with tags to be replaced.
2. `database.json5` saving input numbers and strings.
3. `automate.py` performing calculation and formatting.
### Standardized template
For electricity:
1. Calculate current power draw CPD in kW
2. Calculate proposed power draw PPD in kW
3. Calculate electricity savings ES = CPD * Operating Hours - PPD * Opearting Hours, in kWh/yr
4. Calculate demand savings DS = (CPD - PPD) * Coincidence Factor %/month * 12 months/year, in kW/yr

For natural gas:

5. Calculate natural gas savings NGS (or any other fuel) in MMBtu/yr

Overall:

6. Calculate annnal cost savings ACS = sum(energy savings * unit price), in $

* Use linear equations as much as possible.
* Make sure all physics units can be properly cancelled (also use that as a validation of your template).
### Preparations
1. Make .json5 database by writing detailed comments, including description, unit, data type(int, float, str or list), and default value(if available). The key name could be abbreviation such as "ES", as long as it's consistent with the word document.
2. Clean up word document formatting. In rare scenario, the document could be a legacy .doc file with .docx extension. You need to copy all the text and paste it into a new document.
3. Replace numbers/strings with tags, example: `${XX}`. Make sure to adjust the formatting of the tag, as the format will be preserved.
### Making an automated Python template
1. Read .json5 databases and convert it to `EasyDict`. Then you can easily access the variable by `iac.XX` instead of `iac['XX']`.
2. Perform calculations. Remember to keep the data type consistent which means you'll use `round()` frequently.
3. Format strings. Everything needs to be formatted as strings before replacing. Thousand separator is required. Currency needs to be formatted with $ sign.
4. Import the .docx template.
5. Replace keys with `docx_replace()`.
6. Save file and print caveats if requires more manual operations.
### Equations
Currently, `python-docx-replace` doesn't support replacing keys in Word equations. If possible please use regular linear text instead of equations. If the equation is unavoidable, the workaround is to write the equation in LaTeX then convert it to Word equation and insert it to empty tags like `${XXEqn}`. Check the Reduce Set Pressure template for examples.
### Lookup table
Make a numpy array with table values, then use `np.interp` to get the result.
### Table
If there's any table that needs to be filled with calculated numbers, do the following:

If the table has fixed length, simply access the table and fill in texts.

If the table has variable length, make more reserved lines in the word template, then fill in texts and delete empty lines.

See Repair Leaks template for examples. 
### Blocks
The .docx can have pre-defined blocks with XML tags. E.g. starts with `<XX>` and ends with `</XX>`. Then you may choose to enable/disable the block with `docx_blocks()`