@echo off

REM This script will render all HTML files in the current or Mails directory to PDF using Edge and/or Chrome Browser in the headless mode

setlocal enabledelayedexpansion

set START_DIR=%cd%

REM Check if edge.exe or chrome is installed
set ExeName=
set ExePath=

set EdgeExeName=msedge.exe
set EdgeExePath="C:\Program Files (x86)\Microsoft\Edge\Application\%EdgeExeName%"

set ChromeExeName=chrome.exe
set ChromeExePath="C:\Program Files (x86)\Google\Chrome\Application\%ChromeExeName%"

if exist %EdgeExePath% (

set ExeName=%EdgeExeName%
set ExePath=%EdgeExePath%

) else (
if exist %ChromeExePath% (

set ExeName=%ChromeExeName%
set ExePath=%ChromeExePath%

)
)

if [%ExePath%] == [] (
echo.
echo %EdgeExePath% and %ChromeExePath% browsers don't seem to be installed on this system
echo Please install Edge or Chrome browser and update the path in this script
pause
exit /b
)

if exist index.html (
REM mails exported to separate Html files in Mails directory

if not exist Mails (
echo.
echo Required "Mails" sub-directory doesn't exist, fatal error
echo.
pause
exit /b
)

cd Mails
)

set CURRENT_DIR=%cd%

echo.
echo Using %ExeName% browser in the headless mode to export Mails to PDF
echo.

set HTMLFileName=
set HTMLFileNameBase=

for %%A in ("*.htm") do (

Set HTMLFileName=%%A
Set HTMLFileNameBase=%%~nA

set cmd=call !ExePath!  --headless --disable-gpu --no-pdf-header-footer --print-to-pdf-no-header --print-to-pdf="!CURRENT_DIR!\!HTMLFileNameBase!.pdf" "!CURRENT_DIR!\!HTMLFileName!"
echo !cmd!
!cmd!

echo.
)

cd !START_DIR!
timeout 10
exit /b