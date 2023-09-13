# IAC-Automation
Automated Python assessment for Lehigh University Industrial Assessment Center
##
install VS code, anaconda, go to source control and install git and github desktop.

make a github account w/ lehigh email

sign into github desktop

go to ( https://github.com/BrushXue/IAC-Automation ) on the github website 

click the fork button(make sure both github web and desktop is under your account if you are using the lab computer)

got to github deskop, clone the fork

Inside VS Code, install VS code python plugin, JSON5 syntax

### Set up an environment COPY LINE BY LINE:
```
conda create -n iac python=3.8 
conda activate iac 
```
You can easily set up this env in VSCode.
### Install the following packages:
```
conda install json5 numpy pandas requests
conda install -c conda-forge python-docx latex2mathml num2words
pip install python-docx-replace
```
### To remove this env DO NOT RUN CONDA REMOVE 
```
conda remove --name iac --all
```

## TO COMMIT CHANGES 

make changes in vscode inside the GitHub/... folder

commit changes (github desktop) and push change (github desktop)

go to github website and send a pull request WITH comments of proposed changes (github desktop, branch -> pull request)

Admin can apporove and merge 



## Usage
1. Edit `plant.json5` for general information including energy price
2. Edit any specific `.json5` database
3. Run the corresponding `.py` file
4. The output will be in `ARs` directory

## Semi-automatic AR compiler (beta)
1. Fill in `Info.json5`
2. Copy all AR files into `ARs` directory
3. Run `Compiler.py`
4. Follow the instructions of the script

### Requirements of AR files:
1. No requirement for filename, as long as it's `.docx`
2. Doesn't matter if the file is made from Python template, Excel template, or by hand. The only requirement is no external links (break links if you use Excel template).
3. The title should be "AR *: abcdefg" or "AAR *: abcdefg". Case insensitive. Open View -> Outline, the title should be **level 1**.
5. All sub titles, such as "Recommend Actions", "Anticipated Savings" should be **body text** in outline view. Then set it to **bold, 1.5x line spacing, and 6pt spacing before paragraph**. Otherwise the automatic table of contents will be broken.

## Supported AR templates

### Boiler
Recover Exhaust Gas Heat

### Compressor
Repair Leaks in Compressed Air Lines

### Lighting
Switch to LED lighting

### Others
Install Solar Panel (fully automated using PVWatts API)
