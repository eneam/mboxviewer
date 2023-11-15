@echo off

REM Example script to merge mail files listed in the file.
REM All mail files names must not require Unicode encoding
REM otherwise look at unicode-merge.ps1 Powershell script
REM

set procdir=F:\Documents\GIT1.0.3.40\mboxviewer\x64\Debug


"%procdir%\mboxview.exe" -MBOX_MERGE_LIST_FILE="F:\New1\mboxList2.txt"   -MBOX_MERGE_TO_FILE="F:\NewMerge\mergedFile.mbox"
