@echo off
setlocal EnableExtensions  enabledelayedexpansion

@REM Merging large number of mails into single PDF is very expensive in terms of processing time
@REM This scripts allows users to run HTML to PDF conversion on multiple cores if desired.
@REM Run Print to PDF -> Merge option on large number of selected mails
@REM Before above, Select File->General Options Config->PDF Merge Config and
@REM increase number of HTML encoded mails to merge, to 100 typically or more for simple text mails
@REM to reduce time to convert mails to PDF
@REM Once Mbox Viewer is done with generting HTML files, you can Cancel the Merge processing
@REM and run this script on multiple cores. Copy this scripts and html2pdf-list.cmd script
@REM to directory with generated HTML files
@REM
@REM This main script is similar to run-html2pdf-range.cmd script
@REM This script relies on HTML mail file position instead of HTML mail file names
@REM 
@REM https://okular.kde.org/  PDF viewer capable of viewing very large PDF files
@REM

@echo.
@echo ########################
@echo Running %0 script

set NumberOfBrowsers=4
@REM You must create separate data folder per Browser instance, see UserDataDir below
@REM User Data Folder is not needed if number of browsers is set 1
@REM NumberOfBrowsers windows  will be created and minimized

@REM Script will assume G:\UserAgent\0, G:\UserAgent\1, etc are precreated by user
set RootUserDataDir=G:\UserAgent

SET fileCount=0
FOR /f "tokens=*" %%G IN ('dir /b *.htm') DO ( set /a fileCount+=1 )

@echo off
set TotalFiles=%fileCount%
set /a FilesBlockSize=%TotalFiles%/%NumberOfBrowsers%

@echo.
@echo TotalFiles=%TotalFiles%
@echo FilesBlockSize=%FilesBlockSize%
@echo Number of Browser instances=%NumberOfBrowsers%
@echo.

set /a MaxIndex=NumberOfBrowsers-1

@REM Begin Loop
for /L %%i in (0,1,%MaxIndex%) do (

@set /a FirstFile=%%i*%FilesBlockSize% + 1
@set /a LastFile=%%i*%FilesBlockSize% + %FilesBlockSize%
if %%i == %MaxIndex% ( set LastFile=%TotalFiles% )

set WindowTitle=chrome-list%%i

if "%NumberOfBrowsers%" == "1" ( set UserDataDir= ) ELSE ( set UserDataDir=!RootUserDataDir!\%%i )
if "%UserDataDir%" == "" (
@echo FirstFile=!FirstFile!  LastFile=!LastFile!
) ELSE (
@echo FirstFile=!FirstFile!  LastFile=!LastFile! UserDataDir=!UserDataDir!
)

@echo on
@echo start "!WindowTitle!" /MIN html2pdf-list.cmd !FirstFile! !LastFile! !UserDataDir!
start "!WindowTitle!" /MIN html2pdf-list.cmd !FirstFile! !LastFile! !UserDataDir! ^>%%i.txt 2>&1
@echo off
)
@REM End Loop

pause
exit /b
goto :eof

@REM Run the following to get PIDs of created windows to kill if needed
@REM tasklist /v /FO table /FI "WindowTitle eq chrome-list*"


