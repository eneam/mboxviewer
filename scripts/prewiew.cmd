@echo off

REM Example script to preview first mails in the listed mail files
REM All mail file names must not require Unicode encoding
REM otherwise look at unicode-preview.ps1 Powershell script
REM

set procdir=F:\Documents\GIT1.0.3.40\mboxviewer\x64\Debug

set nameList= ^
	"F:\N\mergetest.mbox" ^
	"F:\New2\message.eml"

	
for %%a in (%nameList%) do (
    "%procdir%\mboxview.exe" -EML_PREVIEW_MODE   -MAIL_FILE="%%a"
)
