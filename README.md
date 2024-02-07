# IAC-Automation
Automated Python script for Lehigh University Industrial Assessment Center

This tutorial is designed for Mechanical Engineering people with zero programming knowledge.
## Required Software
Of course, you need to install Microsoft Office. https://confluence.cc.lehigh.edu/display/LKB/Windows+or+macOS%3A++Download+and+Install+Office+365
### Windows
Install [VS code](https://code.visualstudio.com/download) and [Anaconda](https://www.anaconda.com/download).
### macOS
Install [homebrew](https://brew.sh) then```brew install --cask visual-studio-code anaconda```
### Linux
LibreOffice compatibility is not guaranteed.

## Setting up Python environment
### Open Anaconda Prompt
```
conda create -n iac python=3.8 
conda activate iac 
```
### Install the Following Packages
```
conda install json5 numpy pandas openpyxl requests
conda install -c conda-forge python-docx docxcompose easydict latex2mathml num2words pgeocode
pip install python-docx-replace meteostat
```
`conda` always has the highest priority. If not available, install packages from `conda-forge`. Don't install from `pip` unless you have to, otherwise there might be dependency issue.
### Configure VS Code
Go to `Extensions`(Ctrl+Shift+X), install `Python` (from Microsoft) and `JSON5 syntax` (from mrmlnc).
Press Ctrl+Shift+P, search `Python: Select Interpreter` and select the `iac` environment you just created.
### NOTE: IF YOU WISH TO REMOVE THIS ENVIRONMENT
```
conda remove --name iac --all
```

## Using templates
Is it suggested to work on a **copy** of this reposiotry when generating an IAC report. 

For your convenience, download the main branch from this link: https://codeload.github.com/LehighIAC/IAC-Automation/zip/refs/heads/main , then use VS Code to open the **folder**.

![iac](https://github.com/LehighIAC/IAC-Automation/assets/12702149/fabb6817-7c5a-4e76-9bfe-661a4d2643a5)

### Energy Charts
1. Edit `Energy Charts.xlsx`. Select `fuel type` ,`fuel unit` and `start month`, then edit raw data (if copying from other spreadsheet, copy values only). The formatting is fully automatic and shouldn't be touched.
2. Save the workbook as `Web Page (.htm)` format in the same directory. DO NOT change the filename, all images will be  kept in `Energy Charts.fld` folder. Currently  this is the only stable way to save all charts as images.
3. Run `Utility.py` to extract energy usage data from the spreadsheet.
### Assessment Recommendations
1. Edit `.json5` database of any specific recommendation. Make sure the data type is matching the description.
2. Run the corresponding `.py` file. The output will be saved in `Recommendations` directory. Follow the instructions of the script if there's anything you need to adjust manually.
### Requirements of Manual Recommendation Files:
1. No requirement for filename, as long as it's `.docx`
2. Doesn't matter if the file is made from Python template, Excel template, or by hand. Please **break links** if you used Excel templates.
3. The title text should always be "Recommendation x: Title" or "Additional Recommendation x: Title". Case insensitive. In outline view the title should always be **Level 1**.
4. If there's any other type of energy savings, the unit should be `MMBtu`.
5. All subtitles, such as "Recommend Actions", "Anticipated Savings" should always be **body text** in outline view. Otherwise the automatic table of contents might be broken.
6. In rare cases, some .docx files are actually wrapped legacy .doc file (pre Word 2000) and is not supported by this tool. Please copy and paste everything to a new blank .docx file. 
### Compiling Report
1. Fill required plant information in `Compiler.json5`.
2. Fill other gathered information in `Report/Description.docx`
3. Copy all recommendation documents(if you have any from other sources) into `Recommendations` directory.
4. Run `Compiler.py` to compile the final report.
5. Ctrl+A then F9 to refresh ToC, tables and figures, you need to do it **twice**.

## Supported Recommendation Templates
### Boiler
* Recover Exhaust Gas Heat
* Install Air-Fuel Ratio Controller
### Compressor
* Draw Compressor Intake Air from Outside
* Use Compressor Exhaust to Heat during Winter Months
* Reduce Compressor Set Pressure
* Repair Leaks in Compressed Air Lines
* Install VFD on Existing Air Compressor
* Install New Air Compressor with VFD
### HVAC
* Programmable Thermostat (based on degree hours)
* Replace Old HVAC Units
### Lighting
* Switch to LED lighting (supports any number of areas)
### Motors
* Install VFD on Electric Motor
### Others
* Install Solar Panel (fully automated using PVWatts API)
* Negotiate Utility Charge
