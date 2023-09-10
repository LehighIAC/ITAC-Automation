# IAC-Automation
Automated Python assessment templates

Required Python packages:
```
pip3 install numpy lxml latex2mathml json5 python_docx python_docx_replace num2words requests
```
## Usage
1. Edit `plant.json5` for general information including energy price
2. Edit any specific `.json5` database
3. Run the corresponding `.py` file
4. The output will be in `ARs` directory

## Supported ARs

### Boiler
Recover Exhaust Gas Heat

### Compressor
Repair Leaks in Compressed Air Lines

### Lighting
Switch to LED lighting

### Others
Install Solar Panel (fully automated using PVWatts API)