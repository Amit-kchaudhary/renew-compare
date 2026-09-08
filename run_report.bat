@echo off
cd /d F:\targetpend
call targetpend\Scripts\activate.bat

echo ----------------------------------------------------
echo 1. Generating User Status Summary Report...
echo ----------------------------------------------------
python generate_report.py

echo.
echo ----------------------------------------------------
echo 2. Generating Renewals Comparative Dashboard & Web...
echo ----------------------------------------------------
python generate_renewals_report.py

rem If run with /silent (automated schedule), skip the pause
if "%1"=="/silent" goto end
if "%1"=="--silent" goto end

echo.
echo Process complete. Press any key to close this window.
pause > nul

:end
