@echo off

REM Few hints first :)
REM %~1         - expands %1 removing any surrounding quotes (")
REM %~f1        - expands %1 to a fully qualified path name
REM %~d1        - expands %1 to a drive letter only
REM %~p1        - expands %1 to a path only
REM %~n1        - expands %1 to a file name only
REM %~x1        - expands %1 to a file extension only
REM %~s1        - expanded path contains short names only
REM %~a1        - expands %1 to file attributes
REM %~t1        - expands %1 to date/time of file
REM %~z1        - expands %1 to size of file
REM The modifiers can be combined to get compound results:
REM %~dp1       - expands %1 to a drive letter and path only
REM %~nx1       - expands %1 to a file name and extension only

REM This script is working example script to render all HTML files in the current directory to PDF using wkhtmltopdf tool.
REM
REM wkhtmltopdf.exe version MUST be 0.12.6 or later  !!!!!!!!!!!!!!!!!
REM To verify the version, execute "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" -V

setlocal enabledelayedexpansion

REM Update path if needed
REM Download wkhtmltopdf from https://wkhtmltopdf.org/downloads.html
REM Usage link on how to control header and footer https://wkhtmltopdf.org/usage/wkhtmltopdf.txt

set ProgName=wkhtmltopdf.exe
set ProgDirectoryPath=C:\Program Files\wkhtmltopdf\bin
set CmdPath=%ProgDirectoryPath%\%ProgName%

set NoBackgroundColorOption=--no-background

if NOT exist "%CmdPath%" (

echo.
echo Invalid path to %CmdPath% executable file.
echo Please install and/or update the wkhtmltopdf path in the batch file and re-run the script again.
echo.

pause
exit
)

set HTMLFileName=
set HTMLFileNameBase=

set CURRENT_DIR=%cd%
set PDFDir=%CURRENT_DIR%

for %%A in ("*.htm") do (

Set HTMLFileName=%%A
Set HTMLFileNameBase=%%~nA

REM echo HTMLFileName=!HTMLFileName!
REM echo HTMLFileNameBase=!HTMLFileNameBase!

REM The --no-background option, if present, will remove background color in the target PDF. 

echo "%CmdPath%" --log-level none --zoom 0.9 --enable-local-file-access %NoBackgroundColorOption% --footer-right "Page [page] of [toPage]" "%CURRENT_DIR%\!HTMLFileName!" "%PDFdir%\!HTMLFileNameBase!.pdf" 
call "%CmdPath%" --log-level none --zoom 0.9 --enable-local-file-access %NoBackgroundColorOption% --footer-right "Page [page] of [toPage]" "%CURRENT_DIR%\!HTMLFileName!" "%PDFdir%\!HTMLFileNameBase!.pdf" 

echo.

)

pause