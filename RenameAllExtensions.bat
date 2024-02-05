@echo off
echo ----------------------------------------
echo      Batch File Extension Changer
echo	  USE FOR CONVERTING JPG, PNG ETC TO .1 FILES FOR RELEX
echo
echo      Make sure you select the correct directory before running
echo      
echo ----------------------------------------
echo	  Select Correct Directory
set /p "directory=Enter the directory path: "
echo 	  Enter the old Extension e.g. "jpg, png. &c"
set /p "oldExt=Enter the old extension (without dot): "
echo	  Enter the new Extension e.g. "1" for relex graphics
set /p "newExt=Enter the new extension (without dot): "

cd /d %directory%
ren *.%oldExt% *.%newExt%

echo.
echo File extensions changed successfully.
echo.
pause
