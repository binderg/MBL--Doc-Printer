@echo off
:: Change to the directory containing this batch file
cd /d "%~dp0"

:: Check if the virtual environment exists
if not exist venv\Scripts\activate.bat (
    echo Creating virtual environment...
    call python -m venv venv
)

:: Activate the virtual environment
call venv\Scripts\activate.bat

:: Check if requirements are already installed
python -c "import pkg_resources; exit(0) if pkg_resources.working_set.by_key.get('somepackage') else exit(1)"
if errorlevel 1 (
    echo Installing requirements...
    pip install -r requirements.txt
)

:: Run the Python script
python main.py

:: Close the command prompt window automatically
exit
