# IAC-Automation
Automated Python script for Lehigh University Industrial Assessment Center
## Guide for new IAC members:
0. Of course, you need to have Microsoft Office installed.
   
1. Install [VS code](https://code.visualstudio.com/download), [Anaconda](https://www.anaconda.com/download), [Github Desktop](https://desktop.github.com), and [git](https://gitforwindows.org/) (all default).

2. In VS Code, go to `Extensions`(Ctrl+Shift+X), install `Python` (from Microsoft) and `JSON5 syntax` from mrmlnc.

3. Register a Github account with your lehigh email, then **fork** the [main repository](https://github.com/BrushXue/IAC-Automation). It will make a copy under your account.

4. Sign in Github Desktop, **clone your fork** (not the main repository) to the local computer. The fork should be under your username,

5. After validating your proposed changes, go to Github Desktop to `commit` and `push` new code to your fork. *Remember to write detailed comments so other people can understand your proposed changes.*
   
6. Go to github website (there should be a shortcut in Github Desktop) and send a pull request. *Remeber to write detailed comments so other people can understand your proposed changes.*

7. DO NOT include any sensitive information such as plant name and address when contributing code.

8. After reviewing the code, The IAC Admin can apporove and merge your proposed changes.

## Setting up Python environment:
### Open Anaconda Terminal, or VS Code terminal.
```
conda create -n iac python=3.8 
conda activate iac 
```
### Install the following packages:
```
conda install json5 numpy pandas openpyxl requests
conda install -c conda-forge python-docx docxcompose latex2mathml num2words
pip install python-docx-replace
```
`conda` always has the highest priority. If not available, install packages from `conda-forge`. Don't install from `pip` unless you have to, otherwise there might be dependency issue.
### Config VS Code environment
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

### Requirements of AR files:
1. No requirement for filename, as long as it's `.docx`
2. Doesn't matter if the file is made from Python template, Excel template, or by hand. Please break external links if you used Excel templates.
3. The title text should always be "AR x: Title" or "AAR x: Title". Case insensitive. In outline view the title should always be **level 1**.
4. If there's any other type of energy savings, the unit should be `MMBtu`.
5. All sub titles, such as "Recommend Actions", "Anticipated Savings" should always be **body text** in outline view. Then set it to **bold, 1.5x line spacing, and 6pt spacing before paragraph**. Otherwise the automatic table of contents will be broken.

## Supported AR templates

### Boiler
Recover Exhaust Gas Heat

### Compressor
Repair Leaks in Compressed Air Lines

### Lighting
Switch to LED lighting

### Others
Install Solar Panel (fully automated using PVWatts API)
