@echo off
echo ----------------------------------------
echo      Batch File Extension Changer
echo      Make sure you select the correct directory before running
echo ----------------------------------------
set /p "directory=Enter the directory path: "
set /p "oldExt=Enter the old extension (without dot): "
set /p "newExt=Enter the new extension (without dot): "

cd /d %directory%
ren *.%oldExt% *.%newExt%

echo.
echo File extensions changed successfully.
echo.
pause
