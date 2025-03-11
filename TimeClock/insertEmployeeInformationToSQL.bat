@echo off
setlocal enabledelayedexpansion

:: Define config file location
set CONFIG_FILE=C:\Users\Netadmin\Documents\GitHub\Python\batFileLocations.txt

set SCRIPT_PATH=TimeClock\bat_insertEmployeeInformationToSQL.py

:: Read the anaconda path
for /f "tokens=2 delims==" %%a in ('findstr /I "anaconda=" %CONFIG_FILE%') do (
    set "ANACONDA_PATH=%%a"
)

:: Read the project path
for /f "tokens=2 delims==" %%a in ('findstr /I "project=" %CONFIG_FILE%') do (
    set "PROJECT_PATH=%%a"
)

:: Trim surrounding quotes (if they exist)
set ANACONDA_PATH=%ANACONDA_PATH:"=%
set PROJECT_PATH=%PROJECT_PATH:"=%

:: Ensure the working directory is the Python project root
cd /d "%PROJECT_PATH%"

:: Set PYTHONPATH so Python recognizes this directory as a package root
set PYTHONPATH=%PROJECT_PATH%

:: Debug output (remove these lines once verified)
echo Anaconda Path: "%ANACONDA_PATH%"
echo Project Path: "%PROJECT_PATH%"
echo Python Path: "%PYTHONPATH%"
echo Running: "%ANACONDA_PATH%" "%PROJECT_PATH%\%SCRIPT_PATH%"

:: Run the Python script
"%ANACONDA_PATH%" "%PROJECT_PATH%\%SCRIPT_PATH%"

pause
