@echo on

REM Example script to open selected mail file
REM The selected file name must not require Unicode encoding
REM otherwise look at unicode-open-selected-file.ps1 Powershell script
REM

set procdir=F:\Documents\GIT1.0.3.40\mboxviewer\x64\Debug

set mailFile=F:\N\mergetest.mbox

"%procdir%\mboxview.exe" -MAIL_FILE="%mailFile%"

