@echo off
set /p "directory=Enter the directory path: "
set /p "oldExt=Enter the old extension (without dot): "
set /p "newExt=Enter the new extension (without dot): "

cd /d %directory%
ren *.%oldExt% *.%newExt%

echo File extensions changed successfully.
pause
