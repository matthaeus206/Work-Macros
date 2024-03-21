@echo off

REM Display banner
echo ************************************************************
echo *                 This Script is for adding dates          *
echo *                      to PDFs for Relex                   *
echo *             Dates Must be in (01-01-2024) format         *
echo ************************************************************
echo.

setlocal enabledelayedexpansion

REM Prompt the user to input the directory where the files are located
set /p directory="Enter the directory where the files are located: "

REM Prompt the user to input the text they want to add before the extension
set /p prefix="Enter the text to add before the extension: "

REM Change the current directory to the specified directory
cd /d "%directory%"

REM Check if the directory exists
if not exist "%directory%" (
    echo Directory does not exist.
    pause
    exit /b
)

REM Loop through all .pdf files in the specified directory
for %%F in ("%directory%\*.pdf") do (
    REM Extract the filename without extension
    set "filename=%%~nF"
    
    REM Extract the extension
    set "extension=%%~xF"

    REM Concatenate the prefix, filename, and extension
    set "newfilename=!filename!!prefix!!extension!"

    REM Rename the file
    ren "%%F" "!newfilename!"
)

echo All files renamed successfully.
pause
