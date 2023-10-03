@echo off

echo Starting script...
echo This script will create a virtual environment, install the required packages, and start a terminal session.
echo If you have any issues, please see the README.md file for more information.
echo.

@REM Create a virtual environment named "venv" in the current directory.
echo [1/4] Creating virtual environment...
python -m venv venv

@REM Activate the virtual environment.
echo [2/4] Activating virtual environment...
call venv\Scripts\activate.bat

@REM Install the required packages.
echo [3/4] Installing required packages...
pip install -r requirements.txt

@REM Start a terminal session with the virtual environment activated.
echo [4/4] Starting terminal session...
cmd /k "venv\Scripts\activate.bat"

@REM Pause to see any errors.
echo.
echo Done.
echo.

pause
