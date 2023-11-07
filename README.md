# IAC-Automation
Automated Python script for Lehigh University Industrial Assessment Center
## Guide for New IAC Members
0. Of course, you need to have Microsoft Office installed.
### Windows
1. Install [VS code](https://code.visualstudio.com/download), [Anaconda](https://www.anaconda.com/download), [Github Desktop](https://desktop.github.com), and [git](https://gitforwindows.org/) (all default).
### macOS
0. Install [homebrew](https://brew.sh)

1. ```brew install --cask visual-studio-code anaconda github```

2. In VS Code, go to `Extensions`(Ctrl+Shift+X), install `Python` (from Microsoft) and `JSON5 syntax` from mrmlnc.

3. Register a GitHub account with your Lehigh email, then **fork** the [main repository](https://github.com/LehighIAC/IAC-Automation/tree/main). It will make a copy under your account.

4. Sign in GitHub Desktop, **clone your fork** (not the main repository) to the local computer. The fork should be under your username,

5. After validating your proposed changes, go to Github Desktop to `commit` and `push` new code to your fork. *Remember to write detailed comments so other people can understand your proposed changes.*
   
6. Go to GitHub website (there should be a shortcut in Github Desktop) and send a pull request. *Remember to write detailed comments so other people can understand your proposed changes.*

7. DO NOT include any sensitive information such as plant name and address when contributing code.

8. After reviewing the code, The IAC Admin can approve and merge your proposed changes.

## Setting up Python environment
### Open Anaconda Terminal, or VS Code Terminal.
```
conda create -n iac python=3.8 
conda activate iac 
```
### Install the Following Packages
```
conda install json5 numpy pandas"<2" openpyxl requests
conda install -c conda-forge python-docx docxcompose easydict latex2mathml num2words
pip install python-docx-replace
```
`conda` always has the highest priority. If not available, install packages from `conda-forge`. Don't install from `pip` unless you have to, otherwise there might be dependency issue.
### Configure VS Code Environment
In VS Code, press Ctrl+Shift+P, search `Python: Select Interpreter` and select the `iac` environment you just created.
### NOTE: TO REMOVE THIS ENVIRONMENT
```
conda remove --name iac --all
```
## Usage
Is it suggested to work on a copy of this reposiotry when generating an IAC report.
### Energy Charts
1. Edit `Energy Charts.xlsx`. Select `fuel type` ,`fuel unit` and start month, then edit raw data (if copying from other spreadsheet, copy values only). The formatting is fully automatic and shouldn't be touched.
2. Save the workbook as `Web Page (.htm)` format (DO NOT change the name, all images will be in `Energy Charts.fld` folder). This is the only stable way to save all charts as images.
3. Run `Utility.py` to extract data from the spreadsheet.
### Assessment Recommendations
1. Edit `.json5` database of any specific AR. Make sure the data type is matching the description.
2. Run the corresponding `.py` file. The output will be saved in `ARs` directory. Follow the instructions of the script if there's anything you need to adjust manually.
### Compiling Report
1. Fill plant information in `Info.json5`.
2. Copy all AR files(if you have any from other sources) into `ARs` directory.
3. Run `Compiler.py` to compile the final report.
4. Fill the rest of the information manually.
### Requirements of AR Files:
1. No requirement for filename, as long as it's `.docx`
2. Doesn't matter if the file is made from Python template, Excel template, or by hand. Please break external links if you used Excel templates.
3. The title text should always be "AR x: Title" or "AAR x: Title". Case insensitive. In outline view the title should always be **level 1**.
4. If there's any other type of energy savings, the unit should be `MMBtu`.
5. All sub titles, such as "Recommend Actions", "Anticipated Savings" should always be **body text** in outline view. Then set it to **bold, 1.5x line spacing, and 6pt spacing before paragraph**. Otherwise the automatic table of contents will be broken.

## Template development
**Always make changes in the `develop` branch**!

An automated template is usually made of 3 parts:
1. A .docx template with tags to be replaced.
2. A .json5 database with input numbers and strings.
3. A .py template with calculation and formatting.
### Preparations to make an automated template
1. Make .json5 database by writing detailed comments, including description, unit, data type(int, float, str or list), and default value(if available). The key name could be abbreviation such as "ES", as long as it's consistent with the word document.
2. Clean up word document formatting. In rare scenario, the document could be a legacy .doc file with .docx extension. You need to copy all the text and paste it into a new document.
3. Replace numbers/strings with tags, example: `${ES}`. Make sure to adjust the formatting of the tag, as the format will be preserved.
### Making an automated Python template
1. Read .json5 databases and convert it to `EasyDict`. Then you can easily access the variable by `iac.ES` instead of `iac['ES']`.
2. Perform calculations. Remember to keep the data type consistent which means you'll use `round()` frequently.
3. Format strings. Everything needs to be formatted as strings before replacing. Thousand separator is required. Currency needs to be formatted with $ sign.
4. Import the .docx template.
5. Replace keys with `docx_replace()`.
6. Save file and print caveats if requires manual operations.
### Equations
Currently, `python-docx-replace` doesn't support replacing keys in Word equations. If possible please use regular text instead of equations. If the equation is unavoidable, the workaround is to write the equation in LaTeX then convert it to Word equation and insert it to empty tags like `${ESEqn}`. Check the Lighting template for examples.
### Lookup table
Make a numpy array with table values, then use `np.interp` to get the result.
### Table
If there's any table that needs to be filled with calculated numbers, do the following:

If the table has fixed length, simply access the table and fill in texts.

If the table has variable length, make more reserved lines in the word template, then fill in texts and delete empty lines.

See Repair Leaks template for examples. 
### Blocks
The .docx can have pre-defined blocks with XML tags. E.g. starts with `<AAR>` and ends with `</AAR>`. Then you may choose to enable/disable the block with `docx_blocks()`

## Supported AR Templates

### Boiler
* Recover Exhaust Gas Heat

### Compressor
* Repair Leaks in Compressed Air Lines
* Reduce Compressor Set Pressure

### Lighting
* Switch to LED lighting

### Others
* Install Solar Panel (fully automated using PVWatts API)
