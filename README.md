# IAC-Automation
Automated Python assessment for Lehigh University Industrial Assessment Center
## Guide for new IAC members:
1. When contributing code, DO NOT include any sensitive information such as plant name and address.
   
2. Install [VS code](https://code.visualstudio.com/download), [Anaconda](https://www.anaconda.com/download) and [Github Desktop](https://desktop.github.com).

3. In VS Code, go to `Source Control`(Ctrl+Shift+G) and install git for Windows. goto `Extensions`(Ctrl+Shift+X) and install python, json5.

4. Register a Github account with your lehigh email, then ``fork`` the [main repository](https://github.com/BrushXue/IAC-Automation). It will make a copy under your account.

5. Sign in Github Desktop, `clone` **your fork** (not the main repository) to the local computer. The fork should be under your username,

6. After validating your proposed changes, go to Github Desktop to `commit` and `push` new code to your fork. *Remember to write detailed comments so other people can understand your proposed changes.*
   
7. Go to github website (there should be a shortcut in Github Desktop) and send a pull request. *Remeber to write detailed comments so other people can understand your proposed changes.*

8. After reviewing the code, The IAC Admin can apporove and merge your proposed changes.

## Setting up Python environment:
### Open Anaconda Terminal, or VS Code terminal.
```
conda create -n iac python=3.8 
conda activate iac 
```
### Install the following packages:
```
conda install json5 numpy pandas requests
conda install -c conda-forge python-docx latex2mathml num2words
pip install python-docx-replace
```
### NOTE: TO REMOVE THIS ENVIRONMENT
```
conda remove --name iac --all
```
## Usage
Is it suggested work on a copy of this reposiotry when generating an IAC report.
### Energy Chart
1. Edit `Energy Charts.xlsx`. Select fuel type and unit, then edit raw data (if copying from other spreadsheet, copy values only). The formatting is fully automatic and shouldn't be touched.
2. Click the button to run the macro to save all charts as pictures. You may need to find out how to enable macro on your computer.
3. Run `Utility.py` to extract data from the spreadsheet.
### Recommendations
4. Edit `.json5` database for any specific AR. Make sure the data type is matching the description.
5. Run the corresponding `.py` file. The output will be saved in `ARs` directory. Follow the instructions of the script if there's anything you need to adjust manually.
### Compiling Report
6. Fill plant information in `Info.json5`.
7. Replace plant layout `layout.png`.
8. Copy other AR files(if you made it from other sources) into `ARs` directory.
9.  Run `Compiler.py` to compile the final report.

### Requirements of AR files:
1. No requirement for filename, as long as it's `.docx`
2. Doesn't matter if the file is made from Python template, Excel template, or by hand. The only requirement is no external links (break links if you use Excel template).
3. The title text should always be "AR x: Title" or "AAR x: Title". Case insensitive. Open View -> Outline, the title should always be **level 1**.
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
