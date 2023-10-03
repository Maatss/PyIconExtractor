@echo off

echo Starting script...
echo This script will create a virtual environment, install the required packages, run the script, and then deactivate the virtual environment.
echo If you have any issues, please see the README.md file for more information.
echo.

@REM Create a virtual environment named "venv" in the current directory.
echo [1/5] Creating virtual environment...
python -m venv venv

@REM Activate the virtual environment.
echo [2/5] Activating virtual environment...
call venv\Scripts\activate.bat

@REM Install the required packages.
echo [3/5] Installing required packages...
pip install -r requirements.txt

@REM Run script and pass along all arguments.
echo [4/5] Running script...
python icon_extractor.py %*

@REM Deactivate the virtual environment.
echo [5/5] Deactivating virtual environment...
call venv\Scripts\deactivate.bat

@REM Pause to see any errors.
echo.
echo Done.
echo.

pause
