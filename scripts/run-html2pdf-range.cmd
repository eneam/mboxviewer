@echo off
setlocal enabledelayedexpansion

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
@echo #####################
@echo Running %0 script
@echo.

set NumberOfBrowsers=2
@REM You must create separate data folder per Browser instance, see UserDataDir below
@REM NumberOfBrowsers windows  will be created and minimized

set TotalMails=12
set /a Total=%TotalMails%
set /a MailBlockSize=%Total%/%NumberOfBrowsers%


@echo TotalMails=%TotalMails%
@echo MailBlockSize=%MailBlockSize%
@echo Number of Browser instances=%NumberOfBrowsers%
@echo.

set /a MaxIndex=NumberOfBrowsers-1

@REM Begin Loop
for /L %%i in (0,1,%MaxIndex%) do (

set WindowTitle=chrome-range%%i
set UserDataDir=G:\UserAgent\%%i

@set /a FirstMail=%%i*%MailBlockSize% + 1
@set /a LastMail=%%i*%MailBlockSize% + %MailBlockSize%
if %%i == %MaxIndex% ( set LastMail=%TotalMails% )
@echo FirstMail=!FirstMail!  LastMail=!LastMail! UserDataDir=!UserDataDir!
@echo on
start "!WindowTitle!" /MIN html2pdf-range.cmd !FirstMail! !LastMail! !UserDataDir! ^>%%i.txt 2>&1
@echo off
)
@REM End Loop

goto :eof

@REM Run the following to get PIDs of windows to kill if needed
@REM tasklist /v /FO table /FI "WindowTitle eq chrome-range*"


