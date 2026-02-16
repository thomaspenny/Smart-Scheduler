# TSP Project Launcher - Build Instructions for Single EXE

## Overview
This guide explains how to create a single executable file from the TSP Project Launcher using auto-py-to-exe (PyInstaller).

## Prerequisites

### 1. Install Required Package
```bash
pip install auto-py-to-exe
```

### 2. Verify All Files Are Present
Ensure these files are in the same directory:
- `project_launcher.py` (main file)
- `postcode_distance_app.py`
- `tsp_clustering_app.py`
- `calendar_organizer_app.py`
- `smart_scheduler_app.py`

## Building the EXE

### Step 1: Launch auto-py-to-exe
Open PowerShell in your project directory and run:
```bash
auto-py-to-exe
```

This will open a GUI in your web browser.

### Step 2: Configure Settings

#### Script Location
- **Script Location**: Browse and select `project_launcher.py`

#### Onefile
- **One File**: Select "One File" (this creates a single .exe)

#### Console Window
- **Console Based**: Select "Window Based (hide the console)"

#### Icon (Optional)
- If you have an icon file (.ico), you can select it here

#### Additional Files
Add the Help HTML so the Help button can display it in the EXE:
- Add Data: `help.html` → Destination: `.` (root of the bundled app)

**DO NOT add the Python app files as "Additional Files"** - they will be imported directly into the EXE.

#### Hidden Imports
Click "Add Blank Field" under "Hidden Imports" and add each of these on separate lines:
- `postcode_distance_app`
- `tsp_clustering_app`
- `calendar_organizer_app`
- `smart_scheduler_app`
- `pandas`
- `requests`
- `shapely`
- `sklearn`
- `scipy`
- `matplotlib`
- `win32com.client`
- `win32timezone`

#### Advanced Options
Under "Advanced" tab:
- **Add Data**: Leave empty (no external data files needed)
- **UPX Directory**: Leave empty (unless you have UPX for compression)

### Step 3: Build the EXE

1. Click the big blue "CONVERT .PY TO .EXE" button at the bottom
2. Wait for the build process to complete (may take 1-3 minutes)
3. Check the output folder (default: `output` directory in your project folder)

## Alternative: Using PyInstaller Command Line

If you prefer command line, run this in PowerShell:

```bash
pyinstaller --onefile --windowed --name="TSP_Project_Launcher" --add-data "help.html;." --hidden-import postcode_distance_app --hidden-import tsp_clustering_app --hidden-import calendar_organizer_app --hidden-import smart_scheduler_app --hidden-import pandas --hidden-import requests --hidden-import shapely --hidden-import sklearn --hidden-import scipy --hidden-import matplotlib --hidden-import win32com.client --hidden-import win32timezone -y project_launcher.py
```

**For testing/debugging**, use `--console` instead of `--windowed` to see debug output:
```bash
pyinstaller --onefile --console --name="TSP_Project_Launcher_Debug" --add-data "help.html;." --hidden-import postcode_distance_app --hidden-import tsp_clustering_app --hidden-import calendar_organizer_app --hidden-import smart_scheduler_app --hidden-import pandas --hidden-import requests --hidden-import shapely --hidden-import sklearn --hidden-import scipy --hidden-import matplotlib --hidden-import win32com.client --hidden-import win32timezone -y project_launcher.py
```

**Important**: The app modules MUST be added as `--hidden-import` (not `--add-data`) so PyInstaller properly bundles them into the executable.

## Post-Build Steps

### Test the EXE
1. Navigate to the `output` or `dist` folder
2. Double-click `project_launcher.exe` (or `TSP_Project_Launcher.exe`)
3. Test all four app launches to ensure they work correctly

### Distribution
The single EXE file contains everything needed to run the application:
- All Python code
- All dependencies
- Python interpreter

You can distribute just the .exe file to users who don't have Python installed.

## Troubleshooting

### Error: "Module not found"
- Add the missing module to "Hidden Imports" in auto-py-to-exe
- Common missing modules: `pandas`, `requests`, `win32com`, `pywin32`

### Error: "Failed to launch application"
- Ensure all 5 Python files are in the same directory when building
- Check that all imports at the top of project_launcher.py succeeded

### EXE is very large (>100MB)
- This is normal. PyInstaller bundles the entire Python environment
- Typical size: 100-200MB for apps with pandas, tkinter, etc.

### Apps launch in separate windows instead of as Toplevel
- This is expected behavior when launched as EXE
- Each app runs in its own window for better stability

### Antivirus Flags the EXE
- This is a common false positive with PyInstaller
- You may need to add an exception in your antivirus software
- Digitally signing the EXE can help (requires a code signing certificate)

## File Structure After Build

```
Your Project Folder/
├── project_launcher.py
├── postcode_distance_app.py
├── tsp_clustering_app.py
├── calendar_organizer_app.py
├── smart_scheduler_app.py
├── launcher_config.txt (created on first run)
├── output/ or dist/
│   └── project_launcher.exe  ← Your final EXE
└── build/ (temporary build files - can be deleted)
```

## Notes

### Configuration File
- The `launcher_config.txt` file will be created in the same directory as the EXE on first run
- Settings are now persistent across restarts (config file is saved next to the EXE)
- Project data is stored separately in the configured Projects directory

### Updates
- To update the app, rebuild the EXE with the modified Python files
- User's project data and configuration are preserved (stored separately)

### Dependencies Included
The EXE automatically includes:
- Python 3.x runtime
- tkinter (GUI)
- pandas (data processing)
- requests (API calls for postcodes)
- win32com (Outlook integration, if used)
- All your custom app code

## Recommended Settings Summary

| Setting | Value |
|---------|-------|
| Script Location | project_launcher.py |
| One File | Yes (One File) |
| Console Window | Window Based (hide console) |
| Icon | Optional |
| Additional Files | help.html |
| Hidden Imports | pandas, requests, win32com.client, tkinter |

Good luck with your build! 🚀
