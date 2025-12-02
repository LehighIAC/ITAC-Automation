# IAC-Automation

Automated Python script for Lehigh University Industrial Training and Assessment Center

This tutorial is designed for Mechanical Engineering people with zero programming knowledge.

## Required Software

Install [VS code](https://code.visualstudio.com/download) and [Python](https://www.python.org/downloads/) if you do not already have them.

Of course, you will also need [Microsoft Office](https://lehigh.atlassian.net/wiki/spaces/LKB/pages/26679137/Windows+or+macOS+Download+and+Install+Office+365); LibreOffice compatibility is not guaranteed.

### Configure VS Code

Go to `Extensions`(Ctrl+Shift+X), install `Python` (from Microsoft) and `JSON5 syntax` (from mrmlnc).

### Create Project Virtual Environment and Install Dependencies

1. Using either VS Code's integrated terminal or a separate terminal window, point the terminal to the top-level directory of the project (the directory containing .gitignore, LICENSE, README, setup.py, etc.).
2. Create a virtual environment for the project by running `python -m venv .venv` (or `python3 -m venv .venv`).
3. Activate the virtual environment by running `.venv\\Scripts\\activate` on Windows or `source .venv/bin/activate` on macOS/Linux.
4. With the virtual environment activated, install all of the project's dependencies (listed in setup.py) by running `pip install .` and you should see output as each package is installed.
5. In VS Code, press Ctrl+Shift+P or click on the search bar at the top and select `Show and Run Commands`, search `Python: Select Interpreter` and select the `(.venv)` option corresponding to the virtual environment you just created.
- Note: You can also install individual dependencies to this virtual environment by activating the environment and running the usual `pip install [package]` with any specific package name.
- Note: If you ever need to remove the virtual environment for some reason, you can simply delete the `.venv` directory.

## Using templates

Is it suggested to work on a **copy** of this reposiotry when generating an IAC report.

For your convenience, download the main branch from this link: https://codeload.github.com/LehighIAC/IAC-Automation/zip/refs/heads/main , then use VS Code to open the **folder**.

![iac](https://github.com/LehighIAC/IAC-Automation/assets/12702149/fabb6817-7c5a-4e76-9bfe-661a4d2643a5)

### Energy Charts

1. Edit `Energy Charts.xlsx`. Select `fuel type` ,`fuel unit` and `start month`, then edit raw data (if copying from other spreadsheet, copy values only). The formatting is fully automatic and shouldn't be touched.
2. Save the workbook as `Web Page (.htm)` format in the same directory. DO NOT change the filename, all images will be kept in `Energy Charts.fld` folder. Currently this is the only stable way to save all charts as images.
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

- Install Air-Fuel Ratio Controller [rebate]
- Recover Exhaust Gas Heat [rebate]

### Compressor

- Use Compressor Exhaust to Heat during Winter Months
- Install VFD on Existing Air Compressor [rebate]
- Draw Compressor Intake Air from Outside
- Install New Air Compressor with VFD [rebate]
- Reduce Compressor Set Pressure
- Repair Leaks in Compressed Air Lines

### HVAC

- Install Air Curtain [rebate]
- Insulate Bare Equipemnt
- Programmable Thermostat (based on degree hours)
- Replace Old HVAC Units [rebate]

### Lighting

- Install Motion Sensor [rebate]
- Switch to LED lighting [rebate](supports any number of areas)

### Motors

- Install Industrial Fans to Improve Air Circulation
- Replace Cogged V-Belts
- Install VFD on Electric Motor [rebate]

### Others

- Negotiate Utility Charge
- Install Solar Panel (fully automated using PVWatts API)
