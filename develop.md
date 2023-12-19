# Developing New Templates
### Windows
Install [Github Desktop](https://desktop.github.com), and [git](https://gitforwindows.org/) (all default).
### macOS
```brew install --cask github``` 

1. Register a GitHub account with your Lehigh email, then **fork** the [develop branch](https://github.com/LehighIAC/IAC-Automation/tree/develop). It will make a copy under your account.
2. Sign in GitHub Desktop, clone **your** fork to the local computer. The fork should be under your username.
3. After **validating** your proposed changes, go to Github Desktop to `commit` and `push` new code to your fork.

*Remember to write detailed comments so other people can understand your proposed changes.* 

4. Go to GitHub website (there should be a shortcut in Github Desktop) and send a pull request.

5. DO NOT include any **sensitive** information such as plant name and address when contributing code.
6. After reviewing the code, The IAC Admin can approve and merge your proposed changes.


## Standardized template
An automated template is usually made of 3 parts:
1. `template.docx` with `${tags}` to be replaced.
2. `database.json5` saving input numbers and strings.
3. `automate.py` performing calculation and formatting.

For electricity:
1. Calculate current power draw CPD in [kW]
2. Calculate proposed power draw PPD in [kW]
3. Calculate electricity savings ES = CPD * Operating Hours [hr] - PPD * Opearting Hours [hr], in [kWh/yr]
4. Calculate demand savings DS = (CPD - PPD) * Coincidence Factor [%/month] * n [month/year], in [kW/yr]

For natural gas:
5. Calculate natural gas savings NGS (or any other fuel) in MMBtu/yr

Overall:
6. Calculate annnal cost savings ACS = sum(energy savings * unit price), in $

* Use linear equations as much as possible.
* Make sure all physics units can be properly cancelled (also use that as a validation of your template).
### Preparations
1. Make .json5 database by writing detailed comments, including description, unit, data type(int, float, str or list), and default value (if available). The key name could be abbreviation such as "ES", as long as it's consistent with the word document.
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
